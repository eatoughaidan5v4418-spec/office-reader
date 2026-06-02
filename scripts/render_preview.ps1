param(
    [Parameter(Mandatory = $true)]
    [string]$InputPath,
    [string]$OutputDir = "",
    [int]$TimeoutSeconds = 90,
    [switch]$ContinueAfterComFailure,
    [switch]$ComWorker
)

$ErrorActionPreference = "Stop"
$ScriptPath = $PSCommandPath
$SkillDir = Split-Path -Parent (Split-Path -Parent $ScriptPath)

function Get-PreviewHealthPath {
    if (-not [string]::IsNullOrWhiteSpace($env:OFFICE_READER_PREVIEW_HEALTH_PATH)) {
        return $env:OFFICE_READER_PREVIEW_HEALTH_PATH
    }
    return (Join-Path $SkillDir ".office-reader-cache\preview-backend-health.json")
}

function Assert-PreviewHealthPathSafe {
    $path = Get-PreviewHealthPath
    if ([IO.Path]::GetExtension($path).ToLowerInvariant() -ne ".json") {
        throw "Preview health path must end with .json: $path"
    }
    if (Test-Path -LiteralPath $path -PathType Leaf) {
        try {
            $existing = Get-Content -LiteralPath $path -Raw | ConvertFrom-Json
            if (-not $existing.PSObject.Properties["preview"]) {
                throw "missing preview object"
            }
        } catch {
            throw "Preview health path points to an existing non-health JSON file or non-JSON file: $path"
        }
    }
    return $path
}

function New-PreviewResult {
    param(
        [string]$Status,
        [string]$Backend,
        [string[]]$Artifacts,
        [string[]]$Messages
    )
    [pscustomobject]@{
        status = $Status
        backend = $Backend
        artifacts = $Artifacts
        messages = $Messages
    }
}

function Emit-Failure {
    param([string[]]$Messages)
    New-PreviewResult -Status "failed" -Backend "" -Artifacts @() -Messages $Messages | ConvertTo-Json -Depth 5
    exit 1
}

function Read-PreviewHealth {
    $path = Assert-PreviewHealthPathSafe
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
        return [pscustomobject]@{ preview = [pscustomobject]@{} }
    }
    try {
        return Get-Content -LiteralPath $path -Raw | ConvertFrom-Json
    } catch {
        return [pscustomobject]@{ preview = [pscustomobject]@{} }
    }
}

function Test-OfficeComPreviewUnhealthy {
    param([string]$Extension)
    $health = Read-PreviewHealth
    $preview = $health.preview
    if (-not $preview) { return $false }
    $extensionNode = $preview.PSObject.Properties[$Extension]
    if (-not $extensionNode) { return $false }
    $backendNode = $extensionNode.Value.PSObject.Properties["office-com"]
    if (-not $backendNode) { return $false }
    return ($backendNode.Value.state -eq "unhealthy")
}

function Write-OfficeComPreviewHealth {
    param(
        [string]$Extension,
        [string]$State,
        [string]$Reason,
        [int]$Timeout
    )
    $path = Assert-PreviewHealthPathSafe
    $dir = Split-Path -Parent $path
    if ($dir) {
        New-Item -ItemType Directory -Force -Path $dir | Out-Null
    }
    $health = Read-PreviewHealth
    if (-not $health.preview) {
        $health | Add-Member -NotePropertyName preview -NotePropertyValue ([pscustomobject]@{}) -Force
    }
    if (-not $health.preview.PSObject.Properties[$Extension]) {
        $health.preview | Add-Member -NotePropertyName $Extension -NotePropertyValue ([pscustomobject]@{}) -Force
    }
    $entry = [pscustomobject]@{
        state = $State
        reason = $Reason
        timeout_seconds = $Timeout
        updated_at = (Get-Date).ToUniversalTime().ToString("o")
    }
    $health.preview.$Extension | Add-Member -NotePropertyName "office-com" -NotePropertyValue $entry -Force
    $json = $health | ConvertTo-Json -Depth 8
    $utf8NoBom = [Text.UTF8Encoding]::new($false)
    [IO.File]::WriteAllText($path, $json, $utf8NoBom)
}

function Resolve-LibreOfficeExecutable {
    $command = Get-Command soffice,libreoffice,soffice.exe,libreoffice.exe -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($command -and $command.Source) {
        return $command.Source
    }

    $candidates = @()
    foreach ($value in @($env:LIBREOFFICE_PATH, $env:SOFFICE_PATH)) {
        if ([string]::IsNullOrWhiteSpace($value)) { continue }
        if (Test-Path -LiteralPath $value -PathType Leaf) {
            $candidates += $value
        } elseif (Test-Path -LiteralPath $value -PathType Container) {
            $candidates += (Join-Path $value "soffice.exe")
            $candidates += (Join-Path $value "program\soffice.exe")
        }
    }

    foreach ($root in @($env:ProgramFiles, ${env:ProgramFiles(x86)}, $env:LOCALAPPDATA)) {
        if ([string]::IsNullOrWhiteSpace($root)) { continue }
        $candidates += (Join-Path $root "LibreOffice\program\soffice.exe")
    }

    foreach ($candidate in $candidates) {
        if ($candidate -and (Test-Path -LiteralPath $candidate -PathType Leaf)) {
            return (Resolve-Path -LiteralPath $candidate).Path
        }
    }
    return ""
}

function Stop-OfficeComAutomationProcesses {
    param([ref]$Messages)
    try {
        Get-CimInstance Win32_Process -Filter "Name='WINWORD.EXE' OR Name='POWERPNT.EXE'" -ErrorAction Stop |
            Where-Object { $_.CommandLine -match "/Automation|-Embedding" } |
            ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }
    } catch {
        $Messages.Value += "Could not inspect Office COM automation processes for cleanup: $($_.Exception.Message)"
    }
}

function Invoke-ComWorker {
    param(
        [string]$SourcePath,
        [string]$OutputDirectory,
        [int]$Timeout
    )
    $stdout = Join-Path $OutputDirectory ("preview-worker-" + [guid]::NewGuid().ToString("N") + ".stdout.json")
    $stderr = Join-Path $OutputDirectory ("preview-worker-" + [guid]::NewGuid().ToString("N") + ".stderr.txt")
    $workerScript = if ($ScriptPath) { $ScriptPath } else { $MyInvocation.PSCommandPath }
    if ([string]::IsNullOrWhiteSpace($workerScript)) {
        New-PreviewResult -Status "failed" -Backend "office-com" -Artifacts @() -Messages @("Preview worker script path could not be resolved.") | ConvertTo-Json -Depth 5
        exit 1
    }
    $encodedInput = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($SourcePath))
    $encodedOutput = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($OutputDirectory))
    $encodedScriptPath = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($workerScript))
    $workerCommand = @"
`$scriptPath = [Text.Encoding]::Unicode.GetString([Convert]::FromBase64String('$encodedScriptPath'))
`$inputPath = [Text.Encoding]::Unicode.GetString([Convert]::FromBase64String('$encodedInput'))
`$outputDir = [Text.Encoding]::Unicode.GetString([Convert]::FromBase64String('$encodedOutput'))
& `$scriptPath -InputPath `$inputPath -OutputDir `$outputDir -ComWorker
"@
    $encodedCommand = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($workerCommand))
    $args = @("-NoProfile", "-ExecutionPolicy", "Bypass", "-EncodedCommand", $encodedCommand)
    try {
        $process = Start-Process -FilePath "powershell" -ArgumentList $args -PassThru -WindowStyle Hidden -RedirectStandardOutput $stdout -RedirectStandardError $stderr
    } catch {
        New-PreviewResult -Status "failed" -Backend "office-com" -Artifacts @() -Messages @("Preview worker could not start: $($_.Exception.Message)") | ConvertTo-Json -Depth 5
        exit 1
    }
    if (-not $process.WaitForExit([Math]::Max(1, $Timeout) * 1000)) {
        try { Stop-Process -Id $process.Id -Force -ErrorAction SilentlyContinue } catch {}
        $workerMessages = @()
        Stop-OfficeComAutomationProcesses -Messages ([ref]$workerMessages)
        $message = "Preview rendering timed out after $Timeout seconds."
        if ($workerMessages.Count -gt 0) {
            $message += " " + ($workerMessages -join " ")
        }
        if (Test-Path -LiteralPath $stderr) {
            $errorText = (Get-Content -LiteralPath $stderr -Raw -ErrorAction SilentlyContinue).Trim()
            if ($errorText) { $message += " Worker stderr: $errorText" }
        }
        New-PreviewResult -Status "failed" -Backend "office-com" -Artifacts @() -Messages @($message) | ConvertTo-Json -Depth 5
        exit 1
    }
    $outText = if (Test-Path -LiteralPath $stdout) { Get-Content -LiteralPath $stdout -Raw } else { "" }
    if ([string]::IsNullOrWhiteSpace($outText)) {
        $errText = if (Test-Path -LiteralPath $stderr) { (Get-Content -LiteralPath $stderr -Raw).Trim() } else { "" }
        New-PreviewResult -Status "failed" -Backend "office-com" -Artifacts @() -Messages @("Preview worker exited without JSON output. $errText") | ConvertTo-Json -Depth 5
        exit 1
    }
    Write-Output $outText
    exit $process.ExitCode
}

if (-not (Test-Path -LiteralPath $InputPath -PathType Leaf)) {
    Emit-Failure @("Input file does not exist: $InputPath")
}

$sourcePath = (Resolve-Path -LiteralPath $InputPath).Path
$ext = [IO.Path]::GetExtension($sourcePath).ToLowerInvariant()
if ($ext -notin @(".docx", ".pptx")) {
    Emit-Failure @("Preview rendering expects a normalized .docx or .pptx file; got $ext")
}

$outDir = if ($OutputDir) { $OutputDir } else { Join-Path (Split-Path -Parent $sourcePath) ([IO.Path]::GetFileNameWithoutExtension($sourcePath) + ".preview") }
New-Item -ItemType Directory -Force -Path $outDir | Out-Null
$pdfPath = Join-Path $outDir ([IO.Path]::GetFileNameWithoutExtension($sourcePath) + ".pdf")
if (Test-Path -LiteralPath $pdfPath -PathType Leaf) {
    Emit-Failure @("Preview output already exists and will not be overwritten: $pdfPath")
}

$messages = @()
$skipInlineCom = $false
try {
    Assert-PreviewHealthPathSafe | Out-Null
} catch {
    Emit-Failure @("Invalid preview health path: $($_.Exception.Message)")
}

if (-not $ComWorker) {
    if (Test-OfficeComPreviewUnhealthy -Extension $ext) {
        $messages += "Skipping Office COM preview because backend health memory marks it unhealthy for $ext; trying fallback preview backends first."
        $skipInlineCom = $true
    }
    $hasOfficeCom = (
        ($ext -eq ".docx" -and (Test-Path "Registry::HKEY_CLASSES_ROOT\Word.Application\CLSID")) -or
        ($ext -eq ".pptx" -and (Test-Path "Registry::HKEY_CLASSES_ROOT\PowerPoint.Application\CLSID"))
    )
    if ($hasOfficeCom -and -not $skipInlineCom) {
        if ($ContinueAfterComFailure) {
            $skipInlineCom = $true
            $workerJson = Join-Path $outDir ("preview-worker-main-" + [guid]::NewGuid().ToString("N") + ".json")
            $workerErr = Join-Path $outDir ("preview-worker-main-" + [guid]::NewGuid().ToString("N") + ".err.txt")
            $encodedInput = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($sourcePath))
            $encodedOutput = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($outDir))
            $encodedScriptPath = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($ScriptPath))
            $workerCommand = @"
`$scriptPath = [Text.Encoding]::Unicode.GetString([Convert]::FromBase64String('$encodedScriptPath'))
`$inputPath = [Text.Encoding]::Unicode.GetString([Convert]::FromBase64String('$encodedInput'))
`$outputDir = [Text.Encoding]::Unicode.GetString([Convert]::FromBase64String('$encodedOutput'))
& `$scriptPath -InputPath `$inputPath -OutputDir `$outputDir -ComWorker
"@
            $encodedCommand = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($workerCommand))
            $process = Start-Process -FilePath "powershell" -ArgumentList @("-NoProfile", "-ExecutionPolicy", "Bypass", "-EncodedCommand", $encodedCommand) -PassThru -WindowStyle Hidden -RedirectStandardOutput $workerJson -RedirectStandardError $workerErr
            if (-not $process.WaitForExit([Math]::Max(1, $TimeoutSeconds) * 1000)) {
                try { Stop-Process -Id $process.Id -Force -ErrorAction SilentlyContinue } catch {}
                Stop-OfficeComAutomationProcesses -Messages ([ref]$messages)
                Write-OfficeComPreviewHealth -Extension $ext -State "unhealthy" -Reason "timeout" -Timeout $TimeoutSeconds
                $messages += "Office COM preview timed out after $TimeoutSeconds seconds; continuing to fallback preview backends."
            } else {
                $workerText = if (Test-Path -LiteralPath $workerJson) { Get-Content -LiteralPath $workerJson -Raw } else { "" }
                try {
                    $workerResult = $workerText | ConvertFrom-Json
                    if ($workerResult.status -eq "success") {
                        Write-OfficeComPreviewHealth -Extension $ext -State "healthy" -Reason "success" -Timeout $TimeoutSeconds
                        $workerResult | ConvertTo-Json -Depth 5
                        exit 0
                    }
                    Write-OfficeComPreviewHealth -Extension $ext -State "unhealthy" -Reason "failure" -Timeout $TimeoutSeconds
                    foreach ($message in @($workerResult.messages)) {
                        if ($message) { $messages += "Office COM preview failed before fallback: $message" }
                    }
                } catch {
                    Write-OfficeComPreviewHealth -Extension $ext -State "unhealthy" -Reason "non_json" -Timeout $TimeoutSeconds
                    $messages += "Office COM preview worker returned non-JSON output; continuing to fallback preview backends."
                }
            }
        } else {
            Invoke-ComWorker -SourcePath $sourcePath -OutputDirectory $outDir -Timeout $TimeoutSeconds
        }
    }
}

if (-not $skipInlineCom) {
    try {
        if ($ext -eq ".docx" -and (Test-Path "Registry::HKEY_CLASSES_ROOT\Word.Application\CLSID")) {
            $word = New-Object -ComObject Word.Application
            $word.Visible = $false
            $doc = $word.Documents.Open($sourcePath)
            try {
                $doc.ExportAsFixedFormat($pdfPath, 17)
            } finally {
                $doc.Close([ref]$false)
                $word.Quit()
            }
            if (Test-Path -LiteralPath $pdfPath) {
                Write-OfficeComPreviewHealth -Extension $ext -State "healthy" -Reason "success" -Timeout $TimeoutSeconds
                New-PreviewResult -Status "success" -Backend "office-com" -Artifacts @($pdfPath) -Messages @("Exported DOCX preview PDF with Word COM.") | ConvertTo-Json -Depth 5
                exit 0
            }
        }
    } catch {
        $messages += "Word COM preview failed: $($_.Exception.Message)"
        try { if ($doc) { $doc.Close([ref]$false) }; if ($word) { $word.Quit() } } catch {}
    }

    try {
        if ($ext -eq ".pptx" -and (Test-Path "Registry::HKEY_CLASSES_ROOT\PowerPoint.Application\CLSID")) {
            $powerPoint = New-Object -ComObject PowerPoint.Application
            $presentation = $powerPoint.Presentations.Open($sourcePath, $true, $false, $false)
            try {
                $presentation.SaveAs($pdfPath, 32)
            } finally {
                $presentation.Close()
                $powerPoint.Quit()
            }
            if (Test-Path -LiteralPath $pdfPath) {
                Write-OfficeComPreviewHealth -Extension $ext -State "healthy" -Reason "success" -Timeout $TimeoutSeconds
                New-PreviewResult -Status "success" -Backend "office-com" -Artifacts @($pdfPath) -Messages @("Exported PPTX preview PDF with PowerPoint COM.") | ConvertTo-Json -Depth 5
                exit 0
            }
        }
    } catch {
        $messages += "PowerPoint COM preview failed: $($_.Exception.Message)"
        try { if ($presentation) { $presentation.Close() }; if ($powerPoint) { $powerPoint.Quit() } } catch {}
    }
}

$soffice = Resolve-LibreOfficeExecutable
if ($soffice) {
    $profile = $null
    $libreOfficeSuccess = $false
    try {
        $profile = Join-Path $outDir (".lo_profile_" + [guid]::NewGuid().ToString("N"))
        New-Item -ItemType Directory -Force -Path $profile | Out-Null
        & $soffice --headless "-env:UserInstallation=file:///$($profile.Replace('\','/'))" --convert-to pdf --outdir $outDir $sourcePath | Out-Null
        if (Test-Path -LiteralPath $pdfPath) {
            $libreOfficeSuccess = $true
        }
    } catch {
        $messages += "LibreOffice preview failed: $($_.Exception.Message)"
    } finally {
        if ($profile -and (Test-Path -LiteralPath $profile -PathType Container)) {
            try {
                Remove-Item -LiteralPath $profile -Recurse -Force -ErrorAction Stop
            } catch {
                $messages += "LibreOffice temporary profile cleanup failed: $($_.Exception.Message)"
            }
        }
    }
    if ($libreOfficeSuccess) {
        New-PreviewResult -Status "success" -Backend "libreoffice" -Artifacts @($pdfPath) -Messages @($messages + "Exported preview PDF with LibreOffice.") | ConvertTo-Json -Depth 5
        exit 0
    }
} else {
    $messages += "LibreOffice preview backend not found."
}

if ($messages.Count -eq 0) {
    $messages += "No preview backend was available."
}
Emit-Failure $messages
