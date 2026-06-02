param(
    [Parameter(Mandatory = $true)]
    [string]$InputPath,
    [string]$OutputDir = "",
    [string]$PreferredBackend = "auto",
    [string]$WpsPath = "",
    [string]$WppPath = ""
)

$ErrorActionPreference = "Stop"
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

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

function Save-WithLibreOffice {
    param([string]$Input, [string]$OutputDirectory, [string]$Extension)
    $discoverJson = & powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $ScriptDir "discover_office_backends.ps1") -InputExtension $Extension -Format json | ConvertFrom-Json
    $exe = $discoverJson.backends.libreoffice.executable
    if (-not $exe) { return $false }
    $convertTo = if ($Extension.ToLowerInvariant() -eq ".ppt") { "pptx" } else { "docx" }
    $profile = Join-Path $OutputDirectory (".lo_profile_" + [guid]::NewGuid().ToString("N"))
    try {
        New-Item -ItemType Directory -Force -Path $profile | Out-Null
        & $exe --headless "-env:UserInstallation=file:///$($profile.Replace('\','/'))" --convert-to $convertTo --outdir $OutputDirectory $Input | Out-Null
        return $true
    } finally {
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
        $office = $discovery.backends.office_com
        if (-not $office.available) {
            $messages += "Office COM unavailable: $($office.reason)"
            continue
        }
        $used = Save-WithCom -Input $input -Output $output -Kind $kind -ProgIds @($office.progid)
        if ($used) {
            New-Result -Status "success" -Backend "office-com" -OutputPath $output -Messages @("Converted with $used.") | ConvertTo-Json -Depth 5
            exit 0
        }
        $messages += "Microsoft Office COM was detected, but conversion did not produce output."
    } elseif ($backend -eq "libreoffice") {
        $libre = $discovery.backends.libreoffice
        if (-not $libre.available) {
            $messages += "LibreOffice unavailable: $($libre.reason)"
            continue
        }
        try {
            $ok = Save-WithLibreOffice -Input $input -OutputDirectory $outDir -Extension $ext
            if ($ok -and (Test-Path -LiteralPath $output)) {
                New-Result -Status "success" -Backend "libreoffice" -OutputPath $output -Messages @("Converted with LibreOffice.") | ConvertTo-Json -Depth 5
                exit 0
            }
        } catch {
            $messages += "LibreOffice conversion failed: $($_.Exception.Message)"
        }
    }
}

New-Result -Status "failed" -Backend "" -OutputPath "" -Messages $messages | ConvertTo-Json -Depth 5
exit 1
