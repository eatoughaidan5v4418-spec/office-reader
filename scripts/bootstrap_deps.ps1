param(
    [switch]$DryRun,
    [switch]$IncludeSystemTools,
    [switch]$InstallSystemTools,
    [switch]$IncludeDocling,
    [string]$PythonPath = ""
)

$ErrorActionPreference = "Stop"
$SkillDir = Split-Path -Parent $PSScriptRoot
$VenvDir = Join-Path $SkillDir ".venv"
$VenvPython = Join-Path $VenvDir "Scripts\python.exe"

function Get-LocalToolDirs {
    $dirs = @()
    if ($env:OFFICE_READER_TOOL_PATHS) {
        $dirs += $env:OFFICE_READER_TOOL_PATHS -split [IO.Path]::PathSeparator
    }
    $roots = @(
        (Join-Path $SkillDir "tools"),
        (Split-Path -Parent $SkillDir),
        (Get-Location).Path,
        (Split-Path -Parent (Get-Location).Path)
    )
    foreach ($root in $roots) {
        if (-not (Test-Path -LiteralPath $root)) { continue }
        foreach ($pattern in @("poppler*", "tesseract*", "ocr*", "tools")) {
            $dirs += Get-ChildItem -LiteralPath $root -Directory -Filter $pattern -ErrorAction SilentlyContinue | ForEach-Object { $_.FullName }
        }
    }
    $expanded = @()
    foreach ($dir in $dirs) {
        if ([string]::IsNullOrWhiteSpace($dir)) { continue }
        $expanded += $dir
        $expanded += Join-Path $dir "bin"
        $expanded += Join-Path $dir "Library\bin"
    }
    $expanded | Where-Object { $_ -and (Test-Path -LiteralPath $_) } | Select-Object -Unique
}

function Get-CommandSource {
    param([string[]]$Names)
    foreach ($dir in Get-LocalToolDirs) {
        foreach ($name in $Names) {
            $candidate = Join-Path $dir $name
            if (Test-Path -LiteralPath $candidate -PathType Leaf) {
                return (Resolve-Path -LiteralPath $candidate).Path
            }
        }
    }
    foreach ($name in $Names) {
        $cmd = Get-Command $name -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($cmd) { return $cmd.Source }
    }
    return $null
}

function Test-PythonPackage {
    param([string]$PythonExe, [string]$ImportName)
    if (-not (Test-Path -LiteralPath $PythonExe -PathType Leaf)) { return $false }
    $code = "import importlib.util, sys; sys.exit(0 if importlib.util.find_spec('$ImportName') else 1)"
    & $PythonExe -c $code *> $null
    return ($LASTEXITCODE -eq 0)
}

$pythonPackages = @(
    [pscustomobject]@{ name = "openai"; import_name = "openai"; required = $true },
    [pscustomobject]@{ name = "markitdown-ocr"; import_name = "markitdown_ocr"; required = $true },
    [pscustomobject]@{ name = "rapidocr"; import_name = "rapidocr"; required = $true },
    [pscustomobject]@{ name = "onnxruntime"; import_name = "onnxruntime"; required = $true }
)
if ($IncludeDocling) {
    $pythonPackages += [pscustomobject]@{ name = "docling"; import_name = "docling"; required = $false }
}

$systemTools = @()
if ($IncludeSystemTools -or $InstallSystemTools) {
    $systemTools = @(
        [pscustomobject]@{
            name = "LibreOffice"
            commands = @("soffice", "libreoffice", "soffice.exe", "libreoffice.exe")
            winget_id = "TheDocumentFoundation.LibreOffice"
            choco_id = "libreoffice-fresh"
        },
        [pscustomobject]@{
            name = "Poppler"
            commands = @("pdftoppm")
            winget_id = ""
            choco_id = "poppler"
        },
        [pscustomobject]@{
            name = "Tesseract OCR"
            commands = @("tesseract")
            winget_id = ""
            choco_id = "tesseract"
        },
        [pscustomobject]@{
            name = "WPS Office (optional fallback)"
            commands = @("wps", "wps.exe", "wpp", "wpp.exe")
            winget_id = "Kingsoft.WPSOffice.CN"
            choco_id = "wps-office-free"
            optional = $true
        }
    )
}

$packagePlan = @()
foreach ($pkg in $pythonPackages) {
    $installed = Test-PythonPackage -PythonExe $VenvPython -ImportName $pkg.import_name
    $packagePlan += [pscustomobject]@{
        name = $pkg.name
        import_name = $pkg.import_name
        required = $pkg.required
        installed = $installed
        action = if ($installed) { "none" } else { "pip_install" }
    }
}

$toolPlan = @()
$winget = Get-Command winget -ErrorAction SilentlyContinue
$choco = Get-Command choco -ErrorAction SilentlyContinue
foreach ($tool in $systemTools) {
    $source = Get-CommandSource -Names $tool.commands
    $manager = if ($tool.winget_id -and $winget) { "winget" } elseif ($tool.choco_id -and $choco) { "choco" } else { "" }
    $installCommand = ""
    if (-not $source) {
        if ($manager -eq "winget") {
            $installCommand = "winget install --id $($tool.winget_id) --accept-source-agreements --accept-package-agreements"
        } elseif ($manager -eq "choco") {
            $installCommand = "choco install $($tool.choco_id) -y"
        }
    }
    $toolPlan += [pscustomobject]@{
        name = $tool.name
        available = [bool]$source
        source = $source
        package_manager = $manager
        winget_id = $tool.winget_id
        choco_id = $tool.choco_id
        install_command = $installCommand
        optional = [bool]$tool.optional
        action = if ($source) { "none" } elseif ($installCommand) { "install" } else { "manual_install_required" }
    }
}

if ($DryRun) {
    [pscustomobject]@{
        status = "dry_run"
        skill_dir = $SkillDir
        venv_dir = $VenvDir
        python_packages = $packagePlan
        system_tools = $toolPlan
        messages = @("Dry run only; no packages or system tools were installed.")
    } | ConvertTo-Json -Depth 8
    exit 0
}

if (-not $PythonPath) {
    $codexPython = Join-Path $env:USERPROFILE ".cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
    if (Test-Path -LiteralPath $codexPython -PathType Leaf) {
        $PythonPath = $codexPython
    } else {
        $pythonCommand = Get-Command python -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($pythonCommand) {
            $PythonPath = $pythonCommand.Source
        }
    }
}
if (-not $PythonPath) {
    throw "Python was not found. Pass -PythonPath with a Python 3.10+ executable."
}

if (-not (Test-Path -LiteralPath $VenvPython -PathType Leaf)) {
    & $PythonPath -m venv $VenvDir
    if ($LASTEXITCODE -ne 0) { throw "Failed to create virtual environment at $VenvDir." }
}

& $VenvPython -m pip install --upgrade pip
if ($LASTEXITCODE -ne 0) { throw "Failed to upgrade pip in $VenvDir." }

$toInstall = @($packagePlan | Where-Object { $_.action -eq "pip_install" } | ForEach-Object { $_.name })
if ($toInstall.Count -gt 0) {
    & $VenvPython -m pip install @toInstall
    if ($LASTEXITCODE -ne 0) { throw "Failed to install Python packages: $($toInstall -join ', ')." }
}

$systemMessages = @()
if ($InstallSystemTools) {
foreach ($tool in $toolPlan) {
        if ($tool.action -ne "install") { continue }
        if ($tool.optional) {
            $systemMessages += "Skipped optional fallback tool $($tool.name)."
            continue
        }
        if ($tool.package_manager -eq "winget") {
            winget install --id $tool.winget_id --accept-source-agreements --accept-package-agreements
        } elseif ($tool.package_manager -eq "choco") {
            choco install $tool.choco_id -y
        }
        $systemMessages += "Attempted install for $($tool.name) via $($tool.package_manager)."
    }
} elseif ($toolPlan.Count -gt 0) {
    $systemMessages += "System tools were inspected but not installed. Microsoft Office COM is preferred when available; pass -InstallSystemTools only for fallback tools such as LibreOffice/Poppler/Tesseract or optional WPS."
}

$finalPackagePlan = @()
foreach ($pkg in $pythonPackages) {
    $installed = Test-PythonPackage -PythonExe $VenvPython -ImportName $pkg.import_name
    $finalPackagePlan += [pscustomobject]@{
        name = $pkg.name
        import_name = $pkg.import_name
        required = $pkg.required
        installed = $installed
        action = if ($installed) { "installed" } else { "missing" }
    }
}

[pscustomobject]@{
    status = "completed"
    skill_dir = $SkillDir
    venv_dir = $VenvDir
    python = $VenvPython
    python_packages = $finalPackagePlan
    system_tools = $toolPlan
    messages = $systemMessages
} | ConvertTo-Json -Depth 8
