param(
    [Parameter(Mandatory = $true)]
    [string]$InputPath,
    [string]$OutputPath = ""
)

$ErrorActionPreference = "Stop"

function New-Result {
    param(
        [string]$Status,
        [string]$Backend,
        [string]$OutputPath,
        [string[]]$Messages
    )
    [pscustomobject]@{
        status = $Status
        backend = $Backend
        output_path = $OutputPath
        messages = $Messages
    }
}

if (-not (Test-Path -LiteralPath $InputPath -PathType Leaf)) {
    New-Result -Status "failed" -Backend "" -OutputPath "" -Messages @("Input file does not exist: $InputPath") | ConvertTo-Json -Depth 5
    exit 1
}

$source = (Resolve-Path -LiteralPath $InputPath).Path
$ext = [IO.Path]::GetExtension($source).ToLowerInvariant()
if ($ext -notin @(".doc", ".ppt")) {
    New-Result -Status "failed" -Backend "" -OutputPath "" -Messages @("Legacy text extraction expects .doc or .ppt; got $ext") | ConvertTo-Json -Depth 5
    exit 1
}

if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $OutputPath = Join-Path (Split-Path -Parent $source) ([IO.Path]::GetFileNameWithoutExtension($source) + ".legacy-text.txt")
}

$outDir = Split-Path -Parent $OutputPath
if ($outDir) {
    New-Item -ItemType Directory -Force -Path $outDir | Out-Null
}

try {
    if ($ext -eq ".doc") {
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        $word.DisplayAlerts = 0
        try {
            $doc = $word.Documents.Open($source, $false, $true)
            try {
                $text = $doc.Content.Text
                Set-Content -LiteralPath $OutputPath -Value $text -Encoding UTF8
            } finally {
                if ($doc) { $doc.Close([ref]$false) }
            }
        } finally {
            if ($word) { $word.Quit() }
        }
        New-Result -Status "success" -Backend "word-com-text" -OutputPath $OutputPath -Messages @("Extracted legacy DOC text with Word COM read-only mode.") | ConvertTo-Json -Depth 5
        exit 0
    }

    $powerPoint = New-Object -ComObject PowerPoint.Application
    try {
        $presentation = $powerPoint.Presentations.Open($source, $true, $false, $false)
        $parts = New-Object System.Collections.Generic.List[string]
        try {
            foreach ($slide in $presentation.Slides) {
                $parts.Add("Slide $($slide.SlideIndex)")
                foreach ($shape in $slide.Shapes) {
                    try {
                        if ($shape.HasTextFrame -and $shape.TextFrame.HasText) {
                            $text = $shape.TextFrame.TextRange.Text
                            if (-not [string]::IsNullOrWhiteSpace($text)) {
                                $parts.Add($text)
                            }
                        }
                    } catch {}
                }
                try {
                    if ($slide.NotesPage) {
                        foreach ($shape in $slide.NotesPage.Shapes) {
                            try {
                                if ($shape.HasTextFrame -and $shape.TextFrame.HasText) {
                                    $note = $shape.TextFrame.TextRange.Text
                                    if (-not [string]::IsNullOrWhiteSpace($note)) {
                                        $parts.Add("Notes: $note")
                                    }
                                }
                            } catch {}
                        }
                    }
                } catch {}
            }
            Set-Content -LiteralPath $OutputPath -Value ($parts -join [Environment]::NewLine) -Encoding UTF8
        } finally {
            if ($presentation) { $presentation.Close() }
        }
    } finally {
        if ($powerPoint) { $powerPoint.Quit() }
    }
    New-Result -Status "success" -Backend "powerpoint-com-text" -OutputPath $OutputPath -Messages @("Extracted legacy PPT text with PowerPoint COM read-only mode.") | ConvertTo-Json -Depth 5
    exit 0
} catch {
    try {
        if ($doc) { $doc.Close([ref]$false) }
        if ($word) { $word.Quit() }
        if ($presentation) { $presentation.Close() }
        if ($powerPoint) { $powerPoint.Quit() }
    } catch {}
    New-Result -Status "failed" -Backend "" -OutputPath "" -Messages @("Legacy text extraction failed: $($_.Exception.Message)") | ConvertTo-Json -Depth 5
    exit 1
}
