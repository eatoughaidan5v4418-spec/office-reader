param(
    [string]$InputExtension = ".doc",
    [string]$Format = "json",
    [string]$WpsPath = "",
    [string]$WppPath = ""
)

$ErrorActionPreference = "Stop"
$utf8 = [System.Text.UTF8Encoding]::new($false)
[Console]::OutputEncoding = $utf8
$OutputEncoding = $utf8

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

function Test-ExistingFile {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return $null }
    try {
        $resolved = Resolve-Path -LiteralPath $Path -ErrorAction Stop
        if (Test-Path -LiteralPath $resolved -PathType Leaf) {
            return $resolved.Path
        }
    } catch {
        return $null
    }
    return $null
}

function Get-CommandPath {
    param([string[]]$Names)
    foreach ($name in $Names) {
        $cmd = Get-Command $name -ErrorAction SilentlyContinue
        if ($cmd -and $cmd.Source) { return $cmd.Source }
    }
    return $null
}

function Get-AppPath {
    param([string[]]$ExeNames)
    $roots = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths",
        "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths"
    )
    foreach ($root in $roots) {
        foreach ($exe in $ExeNames) {
            $key = Join-Path $root $exe
            if (Test-Path $key) {
                $value = (Get-ItemProperty $key)."(default)"
                if (-not $value) { $value = (Get-ItemProperty $key).Path }
                $resolved = Test-ExistingFile $value
                if ($resolved) { return $resolved }
            }
        }
    }
    return $null
}

function Find-CommonExe {
    param([string[]]$ExeNames)
    $roots = @(
        "$env:ProgramFiles\WPS Office",
        "${env:ProgramFiles(x86)}\WPS Office",
        "$env:LOCALAPPDATA\Kingsoft\WPS Office",
        "$env:ProgramFiles\Kingsoft\WPS Office",
        "${env:ProgramFiles(x86)}\Kingsoft\WPS Office",
        "$env:ProgramFiles\LibreOffice",
        "${env:ProgramFiles(x86)}\LibreOffice"
    ) | Where-Object { $_ -and (Test-Path $_) }

    foreach ($root in $roots) {
        foreach ($exe in $ExeNames) {
            $found = Get-ChildItem -LiteralPath $root -Recurse -File -Filter $exe -ErrorAction SilentlyContinue |
                Select-Object -First 1
            if ($found) { return $found.FullName }
        }
    }
    return $null
}

function Resolve-LibreOfficeOverride {
    foreach ($value in @($env:SOFFICE_PATH, $env:LIBREOFFICE_PATH)) {
        if ([string]::IsNullOrWhiteSpace($value)) { continue }
        $path = Test-ExistingFile $value
        if ($path) { return $path }
        if (Test-Path -LiteralPath $value -PathType Container) {
            foreach ($candidate in @((Join-Path $value "soffice.exe"), (Join-Path $value "program\soffice.exe"))) {
                $path = Test-ExistingFile $candidate
                if ($path) { return $path }
            }
        }
    }
    return $null
}

function Test-ProgId {
    param([string[]]$ProgIds)
    $available = @()
    foreach ($progId in $ProgIds) {
        if (Test-Path "Registry::HKEY_CLASSES_ROOT\$progId\CLSID") {
            $available += $progId
        }
    }
    return $available
}

function Resolve-WpsBackend {
    param([string]$Ext)
    $isPresentation = $Ext.ToLowerInvariant() -eq ".ppt"
    if ($isPresentation) {
        $override = Test-ExistingFile $WppPath
        if (-not $override) { $override = Test-ExistingFile $env:WPP_PATH }
        if (-not $override) { $override = Test-ExistingFile $env:WPP_EXE }
        $exeNames = @("wpp.exe", "kwpp.exe", "wpsoffice.exe")
        $progIds = @("KWPP.Application", "WPP.Application", "PowerPoint.Application")
    } else {
        $override = Test-ExistingFile $WpsPath
        if (-not $override) { $override = Test-ExistingFile $env:WPS_PATH }
        if (-not $override) { $override = Test-ExistingFile $env:WPS_EXE }
        $exeNames = @("wps.exe", "kwps.exe", "wpsoffice.exe")
        $progIds = @("KWPS.Application", "WPS.Application", "Word.Application")
    }

    $path = $override
    if (-not $path) { $path = Get-CommandPath $exeNames }
    if (-not $path) { $path = Get-AppPath $exeNames }
    if (-not $path) { $path = Find-CommonExe $exeNames }
    $com = Test-ProgId $progIds
    $wpsCom = @($com | Where-Object { $_ -notmatch "^(Word|PowerPoint)\.Application$" })

    [pscustomobject]@{
        name = "wps"
        available = [bool]($path -or $wpsCom)
        executable = $path
        com_progids = $wpsCom
        reason = if ($path) { "WPS executable found" } elseif ($wpsCom) { "WPS-compatible COM ProgID found" } else { "WPS executable or COM ProgID not found" }
    }
}

function Resolve-OfficeComBackend {
    param([string]$Ext)
    $progId = if ($Ext.ToLowerInvariant() -eq ".ppt") { "PowerPoint.Application" } else { "Word.Application" }
    $available = Test-Path "Registry::HKEY_CLASSES_ROOT\$progId\CLSID"
    [pscustomobject]@{
        name = "office-com"
        available = [bool]$available
        progid = $progId
        reason = if ($available) { "Microsoft Office COM ProgID found" } else { "Microsoft Office COM ProgID not found" }
    }
}

function Resolve-LibreOfficeBackend {
    $path = Resolve-LibreOfficeOverride
    if (-not $path) { $path = Get-CommandPath @("soffice", "libreoffice", "soffice.exe", "libreoffice.exe") }
    if (-not $path) { $path = Find-CommonExe @("soffice.exe", "libreoffice.exe") }
    [pscustomobject]@{
        name = "libreoffice"
        available = [bool]$path
        executable = $path
        reason = if ($path) { "LibreOffice executable found" } else { "LibreOffice executable not found" }
    }
}

$ext = $InputExtension.ToLowerInvariant()
if ($ext -notin @(".doc", ".ppt")) {
    throw "InputExtension must be .doc or .ppt; got $InputExtension"
}

$wps = Resolve-WpsBackend $ext
$office = Resolve-OfficeComBackend $ext
$libre = Resolve-LibreOfficeBackend
$priority = @($office, $wps, $libre)

$result = [pscustomobject]@{
    input_extension = $ext
    generated_at = (Get-Date).ToString("o")
    backends = [pscustomobject]@{
        wps = $wps
        office_com = $office
        libreoffice = $libre
    }
    priority_order = $priority
}

if ($Format.ToLowerInvariant() -eq "json") {
    $result | ConvertTo-SafeJson -Depth 8
} else {
    $priority | Format-Table name, available, reason -AutoSize
}
