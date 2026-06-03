param(
    [Parameter(Mandatory = $true)]
    [string]$InputPath,
    [string]$OutputDir = "",
    [string]$PreferredBackend = "auto",
    [string]$WpsPath = "",
    [string]$WppPath = "",
    [int]$TimeoutSeconds = 90,
    [switch]$ContinueAfterComFailure
)

$ErrorActionPreference = "Stop"
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$SkillDir = Split-Path -Parent $ScriptDir

function New-Result {
    param(
        [string]$Status,
        [string]$Backend,
        [string]$OutputPath,
        [string[]]$Messages
    )
    [pscustomobject]@{
        required = $true
        status = $Status
        backend = $Backend
        output_path = $OutputPath
        messages = $Messages
    }
}

function Get-ConversionHealthPath {
    if (-not [string]::IsNullOrWhiteSpace($env:OFFICE_READER_CONVERSION_HEALTH_PATH)) {
        return $env:OFFICE_READER_CONVERSION_HEALTH_PATH
    }
    return (Join-Path $SkillDir ".office-reader-cache\conversion-backend-health.json")
}

function Assert-ConversionHealthPathSafe {
    $path = Get-ConversionHealthPath
    if ([IO.Path]::GetExtension($path).ToLowerInvariant() -ne ".json") {
        throw "Conversion health path must end with .json: $path"
    }
    if (Test-Path -LiteralPath $path -PathType Leaf) {
        try {
            $existing = Get-Content -LiteralPath $path -Raw | ConvertFrom-Json
            if (-not $existing.PSObject.Properties["conversion"]) {
                throw "missing conversion object"
            }
        } catch {
            throw "Conversion health path points to an existing non-health JSON file or non-JSON file: $path"
        }
    }
    return $path
}

function Read-ConversionHealth {
    $path = Assert-ConversionHealthPathSafe
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
        return [pscustomobject]@{ conversion = [pscustomobject]@{} }
    }
    try {
        return Get-Content -LiteralPath $path -Raw | ConvertFrom-Json
    } catch {
        return [pscustomobject]@{ conversion = [pscustomobject]@{} }
    }
}

function Test-OfficeComConversionUnhealthy {
    param([string]$Extension)
    $health = Read-ConversionHealth
    $conversion = $health.conversion
    if (-not $conversion) { return $false }
    $extensionNode = $conversion.PSObject.Properties[$Extension]
    if (-not $extensionNode) { return $false }
    $backendNode = $extensionNode.Value.PSObject.Properties["office-com"]
    if (-not $backendNode) { return $false }
    return ($backendNode.Value.state -eq "unhealthy")
}

function Write-OfficeComConversionHealth {
    param(
        [string]$Extension,
        [string]$State,
        [string]$Reason,
        [int]$Timeout
    )
    $path = Assert-ConversionHealthPathSafe
    $dir = Split-Path -Parent $path
    if ($dir) {
        New-Item -ItemType Directory -Force -Path $dir | Out-Null
    }
    $health = Read-ConversionHealth
    if (-not $health.conversion) {
        $health | Add-Member -NotePropertyName conversion -NotePropertyValue ([pscustomobject]@{}) -Force
    }
    if (-not $health.conversion.PSObject.Properties[$Extension]) {
        $health.conversion | Add-Member -NotePropertyName $Extension -NotePropertyValue ([pscustomobject]@{}) -Force
    }
    $entry = [pscustomobject]@{
        state = $State
        reason = $Reason
        timeout_seconds = $Timeout
        updated_at = (Get-Date).ToUniversalTime().ToString("o")
    }
    $health.conversion.$Extension | Add-Member -NotePropertyName "office-com" -NotePropertyValue $entry -Force
    $json = $health | ConvertTo-Json -Depth 8
    $utf8NoBom = [Text.UTF8Encoding]::new($false)
    [IO.File]::WriteAllText($path, $json, $utf8NoBom)
}

function Save-WithCom {
    param(
        [string]$Input,
        [string]$Output,
        [string]$Kind,
        [string[]]$ProgIds
    )
    foreach ($progId in $ProgIds) {
        try {
            $app = New-Object -ComObject $progId
            $app.Visible = $false
            if ($Kind -eq "word") {
                $doc = $app.Documents.Open($Input)
                try {
                    $doc.SaveAs([ref]$Output, [ref]16)
                } finally {
                    $doc.Close([ref]$false)
                }
            } else {
                $presentation = $app.Presentations.Open($Input, $true, $false, $false)
                try {
                    $presentation.SaveAs($Output, 24)
                } finally {
                    $presentation.Close()
                }
            }
            $app.Quit()
            if (Test-Path -LiteralPath $Output) {
                return $progId
            }
        } catch {
            try {
                if ($doc) { $doc.Close([ref]$false) }
                if ($presentation) { $presentation.Close() }
                if ($app) { $app.Quit() }
            } catch {}
        }
    }
    return $null
}

function Save-WithComWorker {
    param(
        [string]$Input,
        [string]$Output,
        [string]$Kind,
        [string[]]$ProgIds,
        [int]$Timeout
    )
    $ProgIds = @($ProgIds | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    if ($ProgIds.Count -eq 0) {
        return [pscustomobject]@{ status = "failed"; progid = ""; messages = @("Office COM worker had no ProgID to try.") }
    }
    $worker = $env:OFFICE_READER_COM_CONVERSION_WORKER
    $tempWorker = $null
    if ([string]::IsNullOrWhiteSpace($worker)) {
        $tempWorker = Join-Path ([IO.Path]::GetTempPath()) ("office-reader-com-convert-" + [guid]::NewGuid().ToString("N") + ".ps1")
        $workerScript = @'
param([string]$InputPath,[string]$OutputPath,[string]$Kind,[string[]]$ProgIds)
$ErrorActionPreference = "Stop"
foreach ($progId in $ProgIds) {
    $app = $null
    $doc = $null
    $presentation = $null
    try {
        $app = New-Object -ComObject $progId
        $app.Visible = $false
        if ($Kind -eq "word") {
            $doc = $app.Documents.Open($InputPath)
            try {
                $doc.SaveAs([ref]$OutputPath, [ref]16)
            } finally {
                if ($doc) { $doc.Close([ref]$false) }
            }
        } else {
            $presentation = $app.Presentations.Open($InputPath, $true, $false, $false)
            try {
                $presentation.SaveAs($OutputPath, 24)
            } finally {
                if ($presentation) { $presentation.Close() }
            }
        }
        if ($app) { $app.Quit() }
        if (Test-Path -LiteralPath $OutputPath) {
            [pscustomobject]@{ status='success'; progid=$progId } | ConvertTo-Json -Depth 5
            exit 0
        }
    } catch {
        try {
            if ($doc) { $doc.Close([ref]$false) }
            if ($presentation) { $presentation.Close() }
            if ($app) { $app.Quit() }
        } catch {}
    }
}
[pscustomobject]@{ status='failed'; progid='' } | ConvertTo-Json -Depth 5
exit 1
'@
        $utf8NoBom = [Text.UTF8Encoding]::new($false)
        [IO.File]::WriteAllText($tempWorker, $workerScript, $utf8NoBom)
        $worker = $tempWorker
    }
    try {
        $stdout = Join-Path ([IO.Path]::GetTempPath()) ("office-reader-com-convert-" + [guid]::NewGuid().ToString("N") + ".out")
        $stderr = Join-Path ([IO.Path]::GetTempPath()) ("office-reader-com-convert-" + [guid]::NewGuid().ToString("N") + ".err")
        function Quote-Arg([string]$Value) {
            return '"' + ($Value -replace '"', '\"') + '"'
        }
        $workerArgs = @(
            "-NoProfile",
            "-ExecutionPolicy", "Bypass",
            "-File", (Quote-Arg $worker),
            "-InputPath", (Quote-Arg $Input),
            "-OutputPath", (Quote-Arg $Output),
            "-Kind", (Quote-Arg $Kind),
            "-ProgIds"
        )
        $workerArgs += @($ProgIds | ForEach-Object { Quote-Arg ([string]$_) })
        $workerArgLine = ($workerArgs -join " ")
        try {
            $process = Start-Process -FilePath "powershell" -ArgumentList $workerArgLine -WindowStyle Hidden -PassThru -RedirectStandardOutput $stdout -RedirectStandardError $stderr
        } catch {
            return [pscustomobject]@{ status = "failed"; progid = ""; messages = @("Office COM worker could not start: $($_.Exception.Message)") }
        }
        if (-not $process) {
            return [pscustomobject]@{ status = "failed"; progid = ""; messages = @("Office COM worker could not start.") }
        }
        if (-not $process.WaitForExit([Math]::Max(1, $Timeout) * 1000)) {
            try { Stop-Process -Id $process.Id -Force -ErrorAction SilentlyContinue } catch {}
            return [pscustomobject]@{ status = "timeout"; progid = ""; messages = @("Office COM legacy conversion timed out after $Timeout seconds.") }
        }
        $text = if (Test-Path -LiteralPath $stdout) { [string](Get-Content -LiteralPath $stdout -Raw) } else { "" }
        if (-not [string]::IsNullOrWhiteSpace($text)) {
            try {
                $result = $text | ConvertFrom-Json
                if ($result.status -eq "success" -and (Test-Path -LiteralPath $Output)) {
                    return [pscustomobject]@{ status = "success"; progid = $result.progid; messages = @() }
                }
            } catch {}
        }
        return [pscustomobject]@{ status = "failed"; progid = ""; messages = @("Office COM worker did not produce converted output.") }
    } finally {
        foreach ($path in @($stdout, $stderr, $tempWorker)) {
            if ($path -and (Test-Path -LiteralPath $path -PathType Leaf)) {
                try { Remove-Item -LiteralPath $path -Force -ErrorAction SilentlyContinue } catch {}
            }
        }
    }
}

function Save-WithLibreOffice {
    param([string]$Input, [string]$OutputDirectory, [string]$Extension, [int]$Timeout)
    $discoverJson = & powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $ScriptDir "discover_office_backends.ps1") -InputExtension $Extension -Format json | ConvertFrom-Json
    $exe = $discoverJson.backends.libreoffice.executable
    if (-not $exe) { return [pscustomobject]@{ status = "unavailable"; messages = @("LibreOffice executable was not found.") } }
    $convertTo = if ($Extension.ToLowerInvariant() -eq ".ppt") { "pptx" } else { "docx" }
    $profile = Join-Path $OutputDirectory (".lo_profile_" + [guid]::NewGuid().ToString("N"))
    try {
        New-Item -ItemType Directory -Force -Path $profile | Out-Null
        $stdout = Join-Path ([IO.Path]::GetTempPath()) ("office-reader-lo-convert-" + [guid]::NewGuid().ToString("N") + ".out")
        $stderr = Join-Path ([IO.Path]::GetTempPath()) ("office-reader-lo-convert-" + [guid]::NewGuid().ToString("N") + ".err")
        function Quote-Arg([string]$Value) {
            return '"' + ($Value -replace '"', '\"') + '"'
        }
        $loArgs = @(
            "--headless",
            (Quote-Arg "-env:UserInstallation=file:///$($profile.Replace('\','/'))"),
            "--convert-to",
            (Quote-Arg $convertTo),
            "--outdir",
            (Quote-Arg $OutputDirectory),
            (Quote-Arg $Input)
        )
        $loArgLine = ($loArgs -join " ")
        try {
            $psi = New-Object System.Diagnostics.ProcessStartInfo
            $psi.FileName = $exe
            $psi.Arguments = $loArgLine
            $psi.UseShellExecute = $false
            $psi.RedirectStandardOutput = $true
            $psi.RedirectStandardError = $true
            $psi.CreateNoWindow = $true
            $process = New-Object System.Diagnostics.Process
            $process.StartInfo = $psi
            if (-not $process.Start()) {
                return [pscustomobject]@{ status = "failed"; messages = @("LibreOffice conversion could not start.") }
            }
        } catch {
            return [pscustomobject]@{ status = "failed"; messages = @("LibreOffice conversion could not start: $($_.Exception.Message)") }
        }
        if (-not $process.WaitForExit([Math]::Max(1, $Timeout) * 1000)) {
            try { Stop-Process -Id $process.Id -Force -ErrorAction SilentlyContinue } catch {}
            return [pscustomobject]@{ status = "timeout"; messages = @("LibreOffice conversion timed out after $Timeout seconds.") }
        }
        $stdoutText = $process.StandardOutput.ReadToEnd()
        $stderrText = $process.StandardError.ReadToEnd()
        if ($process.ExitCode -eq 0) {
            return [pscustomobject]@{ status = "success"; messages = @() }
        }
        $detail = $stderrText
        if (-not $detail) { $detail = $stdoutText }
        return [pscustomobject]@{ status = "failed"; messages = @("LibreOffice conversion exited with code $($process.ExitCode). $detail") }
    } finally {
        foreach ($path in @($stdout, $stderr)) {
            if ($path -and (Test-Path -LiteralPath $path -PathType Leaf)) {
                try { Remove-Item -LiteralPath $path -Force -ErrorAction SilentlyContinue } catch {}
            }
        }
        if (Test-Path -LiteralPath $profile -PathType Container) {
            try {
                Remove-Item -LiteralPath $profile -Recurse -Force -ErrorAction Stop
            } catch {
                Write-Error -ErrorAction Continue "LibreOffice temporary profile cleanup failed: $($_.Exception.Message)"
            }
        }
    }
}

if (-not (Test-Path -LiteralPath $InputPath -PathType Leaf)) {
    New-Result -Status "failed" -Backend "" -OutputPath "" -Messages @("Input file does not exist: $InputPath") | ConvertTo-Json -Depth 5
    exit 1
}

try {
    Assert-ConversionHealthPathSafe | Out-Null
} catch {
    New-Result -Status "failed" -Backend "" -OutputPath "" -Messages @("Invalid conversion health path: $($_.Exception.Message)") | ConvertTo-Json -Depth 5
    exit 1
}

$input = (Resolve-Path -LiteralPath $InputPath).Path
$ext = [IO.Path]::GetExtension($input).ToLowerInvariant()
if ($ext -notin @(".doc", ".ppt")) {
    New-Result -Status "not_required" -Backend "" -OutputPath $input -Messages @("Input is not .doc or .ppt.") | ConvertTo-Json -Depth 5
    exit 0
}

$outDir = if ($OutputDir) { $OutputDir } else { Split-Path -Parent $input }
New-Item -ItemType Directory -Force -Path $outDir | Out-Null
$targetExt = if ($ext -eq ".doc") { ".docx" } else { ".pptx" }
$output = Join-Path $outDir ([IO.Path]::GetFileNameWithoutExtension($input) + $targetExt)
if (Test-Path -LiteralPath $output -PathType Leaf) {
    New-Result -Status "failed" -Backend "" -OutputPath "" -Messages @("Converted output already exists and will not be overwritten: $output") | ConvertTo-Json -Depth 5
    exit 1
}
$kind = if ($ext -eq ".doc") { "word" } else { "presentation" }

$discoverArgs = @("-NoProfile", "-ExecutionPolicy", "Bypass", "-File", (Join-Path $ScriptDir "discover_office_backends.ps1"), "-InputExtension", $ext, "-Format", "json")
if ($WpsPath) { $discoverArgs += @("-WpsPath", $WpsPath) }
if ($WppPath) { $discoverArgs += @("-WppPath", $WppPath) }
$discovery = & powershell @discoverArgs | ConvertFrom-Json
$messages = @()

$ordered = @("office-com", "wps", "libreoffice")
if ($PreferredBackend -ne "auto") {
    $ordered = @($PreferredBackend) + ($ordered | Where-Object { $_ -ne $PreferredBackend })
}

foreach ($backend in $ordered) {
    if ($backend -eq "wps") {
        $wps = $discovery.backends.wps
        if (-not $wps.available) {
            $messages += "WPS unavailable: $($wps.reason)"
            continue
        }
        $progIds = @($wps.com_progids)
        if ($progIds.Count -eq 0) {
            $progIds = if ($kind -eq "word") { @("KWPS.Application", "WPS.Application") } else { @("KWPP.Application", "WPP.Application") }
        }
        $used = Save-WithCom -Input $input -Output $output -Kind $kind -ProgIds $progIds
        if ($used) {
            New-Result -Status "success" -Backend "wps" -OutputPath $output -Messages @("Converted with $used.") | ConvertTo-Json -Depth 5
            exit 0
        }
        $messages += "WPS was detected, but COM conversion did not produce output. If this WPS build lacks COM automation, use Office COM or LibreOffice."
    } elseif ($backend -eq "office-com") {
        if ((Test-OfficeComConversionUnhealthy -Extension $ext) -and $PreferredBackend -eq "auto") {
            $messages += "Skipping Office COM legacy conversion because backend health memory marks it unhealthy for $ext; trying fallback conversion backends first."
            continue
        }
        $office = $discovery.backends.office_com
        if (-not $office.available) {
            if ([string]::IsNullOrWhiteSpace($env:OFFICE_READER_FAKE_OFFICE_PROGID)) {
                $messages += "Office COM unavailable: $($office.reason)"
                continue
            }
        }
        $officeProgId = if (-not [string]::IsNullOrWhiteSpace($env:OFFICE_READER_FAKE_OFFICE_PROGID)) { $env:OFFICE_READER_FAKE_OFFICE_PROGID } else { $office.progid }
        $workerResult = Save-WithComWorker -Input $input -Output $output -Kind $kind -ProgIds @($officeProgId) -Timeout $TimeoutSeconds
        if ($workerResult.status -eq "success") {
            Write-OfficeComConversionHealth -Extension $ext -State "healthy" -Reason "success" -Timeout $TimeoutSeconds
            New-Result -Status "success" -Backend "office-com" -OutputPath $output -Messages @($messages + "Converted with $($workerResult.progid).") | ConvertTo-Json -Depth 5
            exit 0
        }
        $reason = if ($workerResult.status -eq "timeout") { "timeout" } else { "failure" }
        Write-OfficeComConversionHealth -Extension $ext -State "unhealthy" -Reason $reason -Timeout $TimeoutSeconds
        foreach ($message in @($workerResult.messages)) {
            if ($message) { $messages += $message }
        }
        if ($workerResult.status -ne "timeout") {
            $messages += "Microsoft Office COM was detected, but conversion did not produce output."
        }
        if (-not $ContinueAfterComFailure -and $PreferredBackend -eq "office-com") {
            New-Result -Status "failed" -Backend "office-com" -OutputPath "" -Messages $messages | ConvertTo-Json -Depth 5
            exit 1
        }
    } elseif ($backend -eq "libreoffice") {
        $libre = $discovery.backends.libreoffice
        if (-not $libre.available) {
            $messages += "LibreOffice unavailable: $($libre.reason)"
            continue
        }
        try {
            $loResult = Save-WithLibreOffice -Input $input -OutputDirectory $outDir -Extension $ext -Timeout $TimeoutSeconds
            if ($loResult.status -eq "success" -and (Test-Path -LiteralPath $output)) {
                New-Result -Status "success" -Backend "libreoffice" -OutputPath $output -Messages @($messages + "Converted with LibreOffice.") | ConvertTo-Json -Depth 5
                exit 0
            }
            foreach ($message in @($loResult.messages)) {
                if ($message) { $messages += $message }
            }
        } catch {
            $messages += "LibreOffice conversion failed: $($_.Exception.Message)"
        }
    }
}

New-Result -Status "failed" -Backend "" -OutputPath "" -Messages $messages | ConvertTo-Json -Depth 5
exit 1
