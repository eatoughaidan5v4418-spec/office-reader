param(
    [Parameter(Mandatory = $true)]
    [string]$InputPath,
    [string]$OutputDir = "",
    [string]$PreferredBackend = "auto",
    [string]$WpsPath = "",
    [string]$WppPath = "",
    [ValidateRange(1, 86400)]
    [int]$TimeoutSeconds = 45,
    [switch]$BackendWorker,
    [string]$WorkerBackend = ""
)

$ErrorActionPreference = "Stop"
$utf8 = [System.Text.UTF8Encoding]::new($false)
[Console]::OutputEncoding = $utf8
$OutputEncoding = $utf8
$ScriptPath = $PSCommandPath
$ScriptDir = Split-Path -Parent $ScriptPath

function ConvertTo-SafeJson {
    param(
        [Parameter(ValueFromPipeline = $true)]
        $InputObject,
        [int]$Depth = 5
    )
    process {
        $json = $InputObject | ConvertTo-Json -Depth $Depth
        [regex]::Replace($json, "[^\x00-\x7F]", {
            param($match)
            "\u{0:x4}" -f [int][char]$match.Value
        })
    }
}

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

function Emit-Result {
    param(
        [string]$Status,
        [string]$Backend,
        [string]$OutputPath,
        [string[]]$Messages,
        [int]$ExitCode
    )
    New-Result -Status $Status -Backend $Backend -OutputPath $OutputPath -Messages $Messages |
        ConvertTo-SafeJson -Depth 5
    exit $ExitCode
}

function Get-AutomationProcessIds {
    @(Get-CimInstance Win32_Process -ErrorAction SilentlyContinue |
        Where-Object {
            $_.Name -in @("WINWORD.EXE", "POWERPNT.EXE") -and
            $_.CommandLine -match "/Automation|-Embedding"
        } |
        ForEach-Object { [int]$_.ProcessId })
}

function Stop-NewAutomationProcesses {
    param([int[]]$Before)
    Start-Sleep -Milliseconds 200
    foreach ($processId in Get-AutomationProcessIds) {
        if ($processId -notin $Before) {
            Stop-Process -Id $processId -Force -ErrorAction SilentlyContinue
        }
    }
}

function Save-WithCom {
    param(
        [string]$Input,
        [string]$Output,
        [string]$Kind,
        [string[]]$ProgIds
    )
    if ($env:OFFICE_READER_COM_WORKER_DELAY_SECONDS) {
        Start-Sleep -Seconds ([int]$env:OFFICE_READER_COM_WORKER_DELAY_SECONDS)
    }
    foreach ($progId in $ProgIds) {
        $before = @(Get-AutomationProcessIds)
        $app = $null
        $doc = $null
        $presentation = $null
        try {
            $app = New-Object -ComObject $progId
            if ($Kind -eq "word") {
                $app.Visible = $false
                $app.DisplayAlerts = 0
                $doc = $app.Documents.Open($Input, $false, $true, $false)
                try {
                    $doc.SaveAs2($Output, 16)
                } finally {
                    $doc.Close([ref]$false)
                }
            } else {
                $app.DisplayAlerts = 1
                $presentation = $app.Presentations.Open($Input, $true, $false, $false)
                try {
                    $presentation.SaveAs($Output, 24)
                } finally {
                    $presentation.Close()
                }
            }
            $app.Quit()
            if (Test-Path -LiteralPath $Output -PathType Leaf) {
                return $progId
            }
        } catch {
            try {
                if ($doc) { $doc.Close([ref]$false) }
                if ($presentation) { $presentation.Close() }
                if ($app) { $app.Quit() }
            } catch {}
        } finally {
            Stop-NewAutomationProcesses -Before $before
        }
    }
    return $null
}

function Save-WithLibreOffice {
    param([string]$Input, [string]$OutputDirectory, [string]$Extension)
    $discoverJson = & powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $ScriptDir "discover_office_backends.ps1") -InputExtension $Extension -Format json |
        ConvertFrom-Json
    $exe = $discoverJson.backends.libreoffice.executable
    if (-not $exe) { return $false }
    $convertTo = if ($Extension.ToLowerInvariant() -eq ".ppt") { "pptx" } else { "docx" }
    $profile = Join-Path $OutputDirectory (".lo_profile_" + [guid]::NewGuid().ToString("N"))
    New-Item -ItemType Directory -Force -Path $profile | Out-Null
    try {
        & $exe --headless "-env:UserInstallation=file:///$($profile.Replace('\','/'))" --convert-to $convertTo --outdir $OutputDirectory $Input |
            Out-Null
        return $true
    } finally {
        Remove-LibreOfficeProfile -OutputDirectory $OutputDirectory -ProfilePath $profile
    }
}

function Stop-ProcessTree {
    param([int]$RootProcessId)
    $children = @(Get-CimInstance Win32_Process -Filter "ParentProcessId = $RootProcessId" -ErrorAction SilentlyContinue)
    foreach ($child in $children) {
        Stop-ProcessTree -RootProcessId $child.ProcessId
    }
    Stop-Process -Id $RootProcessId -Force -ErrorAction SilentlyContinue
}

function Remove-WorkerArtifacts {
    param([string[]]$Paths)
    foreach ($path in $Paths) {
        if ($path) {
            for ($attempt = 0; $attempt -lt 3; $attempt++) {
                Remove-Item -LiteralPath $path -Force -ErrorAction SilentlyContinue
                if ($attempt -lt 2) { Start-Sleep -Milliseconds 100 }
            }
        }
    }
}

function Remove-PartialNormalizedOutput {
    param([string]$Path)
    if ($Path) {
        Remove-Item -LiteralPath $Path -Force -ErrorAction SilentlyContinue
    }
}

function Remove-LibreOfficeProfile {
    param([string]$OutputDirectory, [string]$ProfilePath)
    $parent = [IO.Path]::GetFullPath($OutputDirectory).TrimEnd("\") + "\"
    $target = [IO.Path]::GetFullPath($ProfilePath)
    $leaf = Split-Path -Leaf $target
    if ($target.StartsWith($parent, [StringComparison]::OrdinalIgnoreCase) -and $leaf -match "^\.lo_profile_[0-9a-f]+$") {
        Remove-Item -LiteralPath $target -Recurse -Force -ErrorAction SilentlyContinue
    }
}

function Invoke-BackendWorker {
    param(
        [string]$Backend,
        [string]$SourcePath,
        [string]$OutputDirectory,
        [int]$Timeout
    )
    $stdout = Join-Path $OutputDirectory ("legacy-worker-" + [guid]::NewGuid().ToString("N") + ".stdout.json")
    $stderr = Join-Path $OutputDirectory ("legacy-worker-" + [guid]::NewGuid().ToString("N") + ".stderr.txt")
    $encodedScript = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($ScriptPath))
    $encodedInput = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($SourcePath))
    $encodedOutput = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($OutputDirectory))
    $encodedBackend = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($Backend))
    $encodedWps = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($WpsPath))
    $encodedWpp = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($WppPath))
    $workerCommand = @"
`$script = [Text.Encoding]::Unicode.GetString([Convert]::FromBase64String('$encodedScript'))
`$inputPath = [Text.Encoding]::Unicode.GetString([Convert]::FromBase64String('$encodedInput'))
`$outputDir = [Text.Encoding]::Unicode.GetString([Convert]::FromBase64String('$encodedOutput'))
`$backend = [Text.Encoding]::Unicode.GetString([Convert]::FromBase64String('$encodedBackend'))
`$wpsPath = [Text.Encoding]::Unicode.GetString([Convert]::FromBase64String('$encodedWps'))
`$wppPath = [Text.Encoding]::Unicode.GetString([Convert]::FromBase64String('$encodedWpp'))
& `$script -InputPath `$inputPath -OutputDir `$outputDir -WpsPath `$wpsPath -WppPath `$wppPath -BackendWorker -WorkerBackend `$backend
"@
    $encodedCommand = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($workerCommand))
    $automationBefore = @(Get-AutomationProcessIds)
    try {
        $process = Start-Process -FilePath "powershell" `
            -ArgumentList @("-NoProfile", "-ExecutionPolicy", "Bypass", "-EncodedCommand", $encodedCommand) `
            -PassThru -WindowStyle Hidden -RedirectStandardOutput $stdout -RedirectStandardError $stderr
    } catch {
        Remove-WorkerArtifacts -Paths @($stdout, $stderr)
        return [pscustomobject]@{
            status = "failed"
            backend = $Backend
            output_path = ""
            messages = @("$Backend worker could not start: $($_.Exception.Message)")
        }
    }
    if (-not $process.WaitForExit([Math]::Max(1, $Timeout) * 1000)) {
        Stop-ProcessTree -RootProcessId $process.Id
        try { $process.WaitForExit(2000) | Out-Null } catch {}
        Start-Sleep -Milliseconds 200
        Stop-NewAutomationProcesses -Before $automationBefore
        Remove-WorkerArtifacts -Paths @($stdout, $stderr)
        return [pscustomobject]@{
            status = "failed"
            backend = $Backend
            output_path = ""
            messages = @("$Backend conversion timed out after $Timeout seconds; stopped this worker process tree.")
        }
    }
    $text = if (Test-Path -LiteralPath $stdout) { Get-Content -LiteralPath $stdout -Raw } else { "" }
    if ([string]::IsNullOrWhiteSpace($text)) {
        $errorText = if (Test-Path -LiteralPath $stderr) { (Get-Content -LiteralPath $stderr -Raw).Trim() } else { "" }
        Remove-WorkerArtifacts -Paths @($stdout, $stderr)
        return [pscustomobject]@{
            status = "failed"
            backend = $Backend
            output_path = ""
            messages = @("$Backend worker exited without JSON output. $errorText")
        }
    }
    try {
        $result = $text | ConvertFrom-Json
        Remove-WorkerArtifacts -Paths @($stdout, $stderr)
        return $result
    } catch {
        Remove-WorkerArtifacts -Paths @($stdout, $stderr)
        return [pscustomobject]@{
            status = "failed"
            backend = $Backend
            output_path = ""
            messages = @("$Backend worker returned invalid JSON.")
        }
    }
}

if (-not (Test-Path -LiteralPath $InputPath -PathType Leaf)) {
    Emit-Result -Status "failed" -Backend "" -OutputPath "" -Messages @("Input file does not exist: $InputPath") -ExitCode 1
}

$input = (Resolve-Path -LiteralPath $InputPath).Path
$ext = [IO.Path]::GetExtension($input).ToLowerInvariant()
if ($ext -notin @(".doc", ".ppt")) {
    New-Result -Status "not_required" -Backend "" -OutputPath $input -Messages @("Input is not .doc or .ppt.") |
        ConvertTo-SafeJson -Depth 5
    exit 0
}

$outDir = if ($OutputDir) { $OutputDir } else { Split-Path -Parent $input }
New-Item -ItemType Directory -Force -Path $outDir | Out-Null
$targetExt = if ($ext -eq ".doc") { ".docx" } else { ".pptx" }
$baseOutput = Join-Path $outDir ([IO.Path]::GetFileNameWithoutExtension($input) + $targetExt)
$outputCollisionMessage = ""
if (-not $BackendWorker -and (Test-Path -LiteralPath $baseOutput)) {
    $outDir = Join-Path $outDir ("legacy-normalized-" + [guid]::NewGuid().ToString("N"))
    New-Item -ItemType Directory -Force -Path $outDir | Out-Null
    $outputCollisionMessage = "Preserved existing normalized file and converted into a run-specific subdirectory."
}
$output = Join-Path $outDir ([IO.Path]::GetFileNameWithoutExtension($input) + $targetExt)
$kind = if ($ext -eq ".doc") { "word" } else { "presentation" }

$discoverArgs = @("-NoProfile", "-ExecutionPolicy", "Bypass", "-File", (Join-Path $ScriptDir "discover_office_backends.ps1"), "-InputExtension", $ext, "-Format", "json")
if ($WpsPath) { $discoverArgs += @("-WpsPath", $WpsPath) }
if ($WppPath) { $discoverArgs += @("-WppPath", $WppPath) }
$discovery = & powershell @discoverArgs | ConvertFrom-Json

if ($BackendWorker) {
    if ($env:OFFICE_READER_WORKER_PID_LOG) {
        Add-Content -LiteralPath $env:OFFICE_READER_WORKER_PID_LOG -Value "$WorkerBackend,$PID"
    }
    Remove-Item -LiteralPath $output -Force -ErrorAction SilentlyContinue
    if ($WorkerBackend -eq "office-com") {
        $office = $discovery.backends.office_com
        if (-not $office.available) {
            Emit-Result -Status "failed" -Backend "office-com" -OutputPath "" -Messages @("Office COM unavailable: $($office.reason)") -ExitCode 1
        }
        $used = Save-WithCom -Input $input -Output $output -Kind $kind -ProgIds @($office.progid)
        if ($used) {
            Emit-Result -Status "success" -Backend "office-com" -OutputPath $output -Messages @("Converted with $used.") -ExitCode 0
        }
        Emit-Result -Status "failed" -Backend "office-com" -OutputPath "" -Messages @("Microsoft Office COM was detected, but conversion did not produce output.") -ExitCode 1
    }
    if ($WorkerBackend -eq "wps") {
        $wps = $discovery.backends.wps
        if (-not $wps.available) {
            Emit-Result -Status "failed" -Backend "wps" -OutputPath "" -Messages @("WPS unavailable: $($wps.reason)") -ExitCode 1
        }
        $progIds = @($wps.com_progids)
        if ($progIds.Count -eq 0) {
            $progIds = if ($kind -eq "word") { @("KWPS.Application", "WPS.Application") } else { @("KWPP.Application", "WPP.Application") }
        }
        $used = Save-WithCom -Input $input -Output $output -Kind $kind -ProgIds $progIds
        if ($used) {
            Emit-Result -Status "success" -Backend "wps" -OutputPath $output -Messages @("Converted with $used.") -ExitCode 0
        }
        Emit-Result -Status "failed" -Backend "wps" -OutputPath "" -Messages @("WPS was detected, but COM conversion did not produce output.") -ExitCode 1
    }
    if ($WorkerBackend -eq "libreoffice") {
        $libre = $discovery.backends.libreoffice
        if (-not $libre.available) {
            Emit-Result -Status "failed" -Backend "libreoffice" -OutputPath "" -Messages @("LibreOffice unavailable: $($libre.reason)") -ExitCode 1
        }
        try {
            $ok = Save-WithLibreOffice -Input $input -OutputDirectory $outDir -Extension $ext
            if ($ok -and (Test-Path -LiteralPath $output -PathType Leaf)) {
                Emit-Result -Status "success" -Backend "libreoffice" -OutputPath $output -Messages @("Converted with LibreOffice.") -ExitCode 0
            }
        } catch {
            Emit-Result -Status "failed" -Backend "libreoffice" -OutputPath "" -Messages @("LibreOffice conversion failed: $($_.Exception.Message)") -ExitCode 1
        }
        Emit-Result -Status "failed" -Backend "libreoffice" -OutputPath "" -Messages @("LibreOffice conversion did not produce output.") -ExitCode 1
    }
    Emit-Result -Status "failed" -Backend $WorkerBackend -OutputPath "" -Messages @("Unknown worker backend: $WorkerBackend") -ExitCode 1
}

$messages = @()
if ($outputCollisionMessage) {
    $messages += $outputCollisionMessage
}

function Get-DiscoveredBackend {
    param([string]$Backend)
    if ($Backend -eq "office-com") { return $discovery.backends.office_com }
    if ($Backend -eq "wps") { return $discovery.backends.wps }
    if ($Backend -eq "libreoffice") { return $discovery.backends.libreoffice }
    return $null
}

$ordered = @("office-com", "wps", "libreoffice")
if ($PreferredBackend -ne "auto") {
    $ordered = @($PreferredBackend) + ($ordered | Where-Object { $_ -ne $PreferredBackend })
}

foreach ($backend in $ordered) {
    $discoveredBackend = Get-DiscoveredBackend -Backend $backend
    if ($discoveredBackend -and -not $discoveredBackend.available) {
        $messages += "${backend}: Skipping worker because discovery reported unavailable: $($discoveredBackend.reason)"
        continue
    }
    $result = Invoke-BackendWorker -Backend $backend -SourcePath $input -OutputDirectory $outDir -Timeout $TimeoutSeconds
    if ($result.status -eq "success") {
        $messages += @($result.messages)
        Emit-Result -Status "success" -Backend $result.backend -OutputPath $result.output_path -Messages $messages -ExitCode 0
    }
    Remove-PartialNormalizedOutput -Path $output
    foreach ($message in @($result.messages)) {
        if ($message) { $messages += "${backend}: $message" }
    }
}

Remove-PartialNormalizedOutput -Path $output
Emit-Result -Status "failed" -Backend "" -OutputPath "" -Messages $messages -ExitCode 1
