param(
    [Parameter(Mandatory = $true)]
    [string]$InputPath,
    [string]$OutputDir = "",
    [int]$TimeoutSeconds = 90,
    [switch]$ComWorker
)

$ErrorActionPreference = "Stop"
$ScriptPath = $PSCommandPath

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
        Get-CimInstance Win32_Process -Filter "Name='WINWORD.EXE' OR Name='POWERPNT.EXE'" |
            Where-Object { $_.CommandLine -match "/Automation|-Embedding" } |
            ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }
        $message = "Preview rendering timed out after $Timeout seconds."
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

$messages = @()

if (-not $ComWorker) {
    $hasOfficeCom = (
        ($ext -eq ".docx" -and (Test-Path "Registry::HKEY_CLASSES_ROOT\Word.Application\CLSID")) -or
        ($ext -eq ".pptx" -and (Test-Path "Registry::HKEY_CLASSES_ROOT\PowerPoint.Application\CLSID"))
    )
    if ($hasOfficeCom) {
        Invoke-ComWorker -SourcePath $sourcePath -OutputDirectory $outDir -Timeout $TimeoutSeconds
    }
}

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
            New-PreviewResult -Status "success" -Backend "office-com" -Artifacts @($pdfPath) -Messages @("Exported PPTX preview PDF with PowerPoint COM.") | ConvertTo-Json -Depth 5
            exit 0
        }
    }
} catch {
    $messages += "PowerPoint COM preview failed: $($_.Exception.Message)"
    try { if ($presentation) { $presentation.Close() }; if ($powerPoint) { $powerPoint.Quit() } } catch {}
}

$soffice = Get-Command soffice,libreoffice,soffice.exe,libreoffice.exe -ErrorAction SilentlyContinue | Select-Object -First 1
if ($soffice) {
    try {
        $profile = Join-Path $outDir (".lo_profile_" + [guid]::NewGuid().ToString("N"))
        New-Item -ItemType Directory -Force -Path $profile | Out-Null
        & $soffice.Source --headless "-env:UserInstallation=file:///$($profile.Replace('\','/'))" --convert-to pdf --outdir $outDir $sourcePath | Out-Null
        if (Test-Path -LiteralPath $pdfPath) {
            New-PreviewResult -Status "success" -Backend "libreoffice" -Artifacts @($pdfPath) -Messages @("Exported preview PDF with LibreOffice.") | ConvertTo-Json -Depth 5
            exit 0
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
Emit-Failure $messages
