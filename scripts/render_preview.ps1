param(
    [Parameter(Mandatory = $true)]
    [string]$InputPath,
    [string]$OutputDir = "",
    [ValidateRange(1, 86400)]
    [int]$TimeoutSeconds = 90,
    [switch]$ContinueAfterComFailure,
    [switch]$ComWorker
)

$ErrorActionPreference = "Stop"
$utf8 = [System.Text.UTF8Encoding]::new($false)
[Console]::OutputEncoding = $utf8
$OutputEncoding = $utf8
$ScriptPath = $PSCommandPath

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
    New-PreviewResult -Status "failed" -Backend "" -Artifacts @() -Messages $Messages | ConvertTo-SafeJson -Depth 5
    exit 1
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

function Remove-PartialPreview {
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

function Get-PreviewHealthPath {
    if (-not [string]::IsNullOrWhiteSpace($env:OFFICE_READER_PREVIEW_HEALTH_PATH)) {
        return $env:OFFICE_READER_PREVIEW_HEALTH_PATH
    }
    $root = if (-not [string]::IsNullOrWhiteSpace($env:LOCALAPPDATA)) {
        $env:LOCALAPPDATA
    } else {
        [IO.Path]::GetTempPath()
    }
    Join-Path $root "office-reader\preview-backend-health.json"
}

function Read-PreviewHealth {
    $path = Get-PreviewHealthPath
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
        return [pscustomobject]@{}
    }
    try {
        $health = Get-Content -LiteralPath $path -Raw -Encoding UTF8 | ConvertFrom-Json
        if ($health) { return $health }
    } catch {}
    [pscustomobject]@{}
}

function Write-PreviewHealth {
    param($Health)
    $path = Get-PreviewHealthPath
    try {
        $parent = Split-Path -Parent $path
        if ($parent) { New-Item -ItemType Directory -Force -Path $parent | Out-Null }
        $temporary = "$path.$([guid]::NewGuid().ToString('N')).tmp"
        try {
            $Health | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $temporary -Encoding UTF8
            Move-Item -LiteralPath $temporary -Destination $path -Force
        } finally {
            Remove-Item -LiteralPath $temporary -Force -ErrorAction SilentlyContinue
        }
    } catch {}
}

function Set-PreviewComDegraded {
    param([string]$Extension, [string]$Reason)
    if ($Extension -ne ".docx") { return }
    $health = Read-PreviewHealth
    $health | Add-Member -NotePropertyName "word-com-preview" -NotePropertyValue ([pscustomobject]@{
        prefer_libreoffice_until_utc = [DateTime]::UtcNow.AddDays(7).ToString("o")
        reason = $Reason
        updated_utc = [DateTime]::UtcNow.ToString("o")
    }) -Force
    Write-PreviewHealth -Health $health
}

function Clear-PreviewComDegraded {
    param([string]$Extension)
    if ($Extension -ne ".docx") { return }
    $health = Read-PreviewHealth
    if ($health.PSObject.Properties.Name -contains "word-com-preview") {
        $health.PSObject.Properties.Remove("word-com-preview")
        Write-PreviewHealth -Health $health
    }
}

function Test-PreferLibreOfficePreview {
    param([string]$Extension)
    if ($Extension -ne ".docx") { return $false }
    try {
        $entry = (Read-PreviewHealth)."word-com-preview"
        if (-not $entry.prefer_libreoffice_until_utc) { return $false }
        return [DateTime]::Parse($entry.prefer_libreoffice_until_utc).ToUniversalTime() -gt [DateTime]::UtcNow
    } catch {
        return $false
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
        New-PreviewResult -Status "failed" -Backend "office-com" -Artifacts @() -Messages @("Preview worker script path could not be resolved.") | ConvertTo-SafeJson -Depth 5
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
    $automationBefore = @(Get-AutomationProcessIds)
    try {
        $process = Start-Process -FilePath "powershell" -ArgumentList $args -PassThru -WindowStyle Hidden -RedirectStandardOutput $stdout -RedirectStandardError $stderr
    } catch {
        Remove-WorkerArtifacts -Paths @($stdout, $stderr)
        New-PreviewResult -Status "failed" -Backend "office-com" -Artifacts @() -Messages @("Preview worker could not start: $($_.Exception.Message)") | ConvertTo-SafeJson -Depth 5
        exit 1
    }
    if (-not $process.WaitForExit([Math]::Max(1, $Timeout) * 1000)) {
        Stop-ProcessTree -RootProcessId $process.Id
        try { $process.WaitForExit(2000) | Out-Null } catch {}
        Start-Sleep -Milliseconds 200
        Stop-NewAutomationProcesses -Before $automationBefore
        Set-PreviewComDegraded -Extension $ext -Reason "timeout"
        $message = "Preview rendering timed out after $Timeout seconds."
        if (Test-Path -LiteralPath $stderr) {
            $errorText = (Get-Content -LiteralPath $stderr -Raw -ErrorAction SilentlyContinue).Trim()
            if ($errorText) { $message += " Worker stderr: $errorText" }
        }
        Remove-WorkerArtifacts -Paths @($stdout, $stderr)
        New-PreviewResult -Status "failed" -Backend "office-com" -Artifacts @() -Messages @($message) | ConvertTo-SafeJson -Depth 5
        exit 1
    }
    $outText = if (Test-Path -LiteralPath $stdout) { Get-Content -LiteralPath $stdout -Raw } else { "" }
    if ([string]::IsNullOrWhiteSpace($outText)) {
        $errText = if (Test-Path -LiteralPath $stderr) { (Get-Content -LiteralPath $stderr -Raw).Trim() } else { "" }
        Remove-WorkerArtifacts -Paths @($stdout, $stderr)
        New-PreviewResult -Status "failed" -Backend "office-com" -Artifacts @() -Messages @("Preview worker exited without JSON output. $errText") | ConvertTo-SafeJson -Depth 5
        exit 1
    }
    Remove-WorkerArtifacts -Paths @($stdout, $stderr)
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
$basePdfPath = Join-Path $outDir ([IO.Path]::GetFileNameWithoutExtension($sourcePath) + ".pdf")
$outputCollisionMessage = ""
if (-not $ComWorker -and (Test-Path -LiteralPath $basePdfPath)) {
    $outDir = Join-Path $outDir ("preview-render-" + [guid]::NewGuid().ToString("N"))
    New-Item -ItemType Directory -Force -Path $outDir | Out-Null
    $outputCollisionMessage = "Preserved existing preview PDF and rendered into a run-specific subdirectory."
}
$pdfPath = Join-Path $outDir ([IO.Path]::GetFileNameWithoutExtension($sourcePath) + ".pdf")

$messages = @()
if ($outputCollisionMessage) {
    $messages += $outputCollisionMessage
}
$skipInlineCom = $false
$preferLibreOffice = (-not $ComWorker) -and (Test-PreferLibreOfficePreview -Extension $ext)

if ($preferLibreOffice) {
    $messages += "Preview backend health memory prefers LibreOffice before Word COM after a recent Word COM timeout."
    Remove-PartialPreview -Path $pdfPath
    $preferredSoffice = Resolve-LibreOfficeExecutable
    if ($preferredSoffice) {
        try {
            $preferredProfile = Join-Path $outDir (".lo_profile_" + [guid]::NewGuid().ToString("N"))
            New-Item -ItemType Directory -Force -Path $preferredProfile | Out-Null
            try {
                & $preferredSoffice --headless "-env:UserInstallation=file:///$($preferredProfile.Replace('\','/'))" --convert-to pdf --outdir $outDir $sourcePath | Out-Null
                if (Test-Path -LiteralPath $pdfPath) {
                    New-PreviewResult -Status "success" -Backend "libreoffice" -Artifacts @($pdfPath) -Messages @($messages + "Exported preview PDF with preferred LibreOffice backend.") | ConvertTo-SafeJson -Depth 5
                    exit 0
                }
            } finally {
                Remove-LibreOfficeProfile -OutputDirectory $outDir -ProfilePath $preferredProfile
            }
        } catch {
            $messages += "Preferred LibreOffice preview failed: $($_.Exception.Message)"
        }
    } else {
        $messages += "Preferred LibreOffice preview backend was not found; retrying the normal preview order."
    }
    Remove-PartialPreview -Path $pdfPath
}

if (-not $ComWorker) {
    $hasOfficeCom = (
        ($ext -eq ".docx" -and (Test-Path "Registry::HKEY_CLASSES_ROOT\Word.Application\CLSID")) -or
        ($ext -eq ".pptx" -and (Test-Path "Registry::HKEY_CLASSES_ROOT\PowerPoint.Application\CLSID"))
    )
    if ($hasOfficeCom) {
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
            $automationBefore = @(Get-AutomationProcessIds)
            $process = Start-Process -FilePath "powershell" -ArgumentList @("-NoProfile", "-ExecutionPolicy", "Bypass", "-EncodedCommand", $encodedCommand) -PassThru -WindowStyle Hidden -RedirectStandardOutput $workerJson -RedirectStandardError $workerErr
            if (-not $process.WaitForExit([Math]::Max(1, $TimeoutSeconds) * 1000)) {
                Stop-ProcessTree -RootProcessId $process.Id
                try { $process.WaitForExit(2000) | Out-Null } catch {}
                Start-Sleep -Milliseconds 200
                Stop-NewAutomationProcesses -Before $automationBefore
                Set-PreviewComDegraded -Extension $ext -Reason "timeout"
                $messages += "Office COM preview timed out after $TimeoutSeconds seconds; continuing to fallback preview backends."
            } else {
                $workerText = if (Test-Path -LiteralPath $workerJson) { Get-Content -LiteralPath $workerJson -Raw } else { "" }
                try {
                    $workerResult = $workerText | ConvertFrom-Json
                    if ($workerResult.status -eq "success") {
                        Clear-PreviewComDegraded -Extension $ext
                        if ($outputCollisionMessage) {
                            $workerResult.messages = @($outputCollisionMessage) + @($workerResult.messages)
                        }
                        Remove-WorkerArtifacts -Paths @($workerJson, $workerErr)
                        $workerResult | ConvertTo-SafeJson -Depth 5
                        exit 0
                    }
                    foreach ($message in @($workerResult.messages)) {
                        if ($message) { $messages += "Office COM preview failed before fallback: $message" }
                    }
                } catch {
                    $messages += "Office COM preview worker returned non-JSON output; continuing to fallback preview backends."
                }
            }
            Remove-WorkerArtifacts -Paths @($workerJson, $workerErr)
        } else {
            Invoke-ComWorker -SourcePath $sourcePath -OutputDirectory $outDir -Timeout $TimeoutSeconds
        }
    }
}

if (-not $skipInlineCom) {
    $trackAutomation = $false
    try {
        if ($ext -eq ".docx" -and (Test-Path "Registry::HKEY_CLASSES_ROOT\Word.Application\CLSID")) {
            $before = @(Get-AutomationProcessIds)
            $trackAutomation = $true
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
                Clear-PreviewComDegraded -Extension $ext
                New-PreviewResult -Status "success" -Backend "office-com" -Artifacts @($pdfPath) -Messages @($messages + "Exported DOCX preview PDF with Word COM.") | ConvertTo-SafeJson -Depth 5
                exit 0
            }
        }
    } catch {
        $messages += "Word COM preview failed: $($_.Exception.Message)"
        try { if ($doc) { $doc.Close([ref]$false) }; if ($word) { $word.Quit() } } catch {}
    } finally {
        if ($trackAutomation) { Stop-NewAutomationProcesses -Before $before }
    }

    $trackAutomation = $false
    try {
        if ($ext -eq ".pptx" -and (Test-Path "Registry::HKEY_CLASSES_ROOT\PowerPoint.Application\CLSID")) {
            $before = @(Get-AutomationProcessIds)
            $trackAutomation = $true
            $powerPoint = New-Object -ComObject PowerPoint.Application
            $presentation = $powerPoint.Presentations.Open($sourcePath, $true, $false, $false)
            try {
                $presentation.SaveAs($pdfPath, 32)
            } finally {
                $presentation.Close()
                $powerPoint.Quit()
            }
            if (Test-Path -LiteralPath $pdfPath) {
                New-PreviewResult -Status "success" -Backend "office-com" -Artifacts @($pdfPath) -Messages @($messages + "Exported PPTX preview PDF with PowerPoint COM.") | ConvertTo-SafeJson -Depth 5
                exit 0
            }
        }
    } catch {
        $messages += "PowerPoint COM preview failed: $($_.Exception.Message)"
        try { if ($presentation) { $presentation.Close() }; if ($powerPoint) { $powerPoint.Quit() } } catch {}
    } finally {
        if ($trackAutomation) { Stop-NewAutomationProcesses -Before $before }
    }
}

Remove-PartialPreview -Path $pdfPath
$soffice = Resolve-LibreOfficeExecutable
if ($soffice) {
    try {
        $profile = Join-Path $outDir (".lo_profile_" + [guid]::NewGuid().ToString("N"))
        New-Item -ItemType Directory -Force -Path $profile | Out-Null
        try {
            & $soffice --headless "-env:UserInstallation=file:///$($profile.Replace('\','/'))" --convert-to pdf --outdir $outDir $sourcePath | Out-Null
            if (Test-Path -LiteralPath $pdfPath) {
                New-PreviewResult -Status "success" -Backend "libreoffice" -Artifacts @($pdfPath) -Messages @($messages + "Exported preview PDF with LibreOffice.") | ConvertTo-SafeJson -Depth 5
                exit 0
            }
        } finally {
            Remove-LibreOfficeProfile -OutputDirectory $outDir -ProfilePath $profile
        }
    } catch {
        $messages += "LibreOffice preview failed: $($_.Exception.Message)"
    }
} else {
    $messages += "LibreOffice preview backend not found."
}

if ($messages.Count -eq 0) {
    $messages += "No preview backend was available."
}
Remove-PartialPreview -Path $pdfPath
Emit-Failure $messages
