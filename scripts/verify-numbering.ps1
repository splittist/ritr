param([string]$Expected = 'test-results/numbering-expected.json')

$ErrorActionPreference = 'Stop'
$cases = Get-Content -LiteralPath $Expected -Raw -Encoding utf8 | ConvertFrom-Json
$word = $null
$failures = [System.Collections.Generic.List[string]]::new()
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $word.AutomationSecurity = 3
    foreach ($case in $cases) {
        $document = $null
        try {
            $document = $word.Documents.Open($case.file, $false, $true, $false)
            if ($document.Paragraphs.Count -ne $case.paragraphs.Count) { throw "Paragraph count mismatch: $($case.file)" }
            for ($i = 0; $i -lt $case.paragraphs.Count; $i++) {
                $actual = $document.Paragraphs.Item($i + 1)
                $expectedParagraph = $case.paragraphs[$i]
                $label = $actual.Range.ListFormat.ListString
                if ($label -cne $expectedParagraph.label) { $failures.Add("Label mismatch at paragraph $($i + 1) in $($case.file): Word='$label', ritr='$($expectedParagraph.label)'") }
                foreach ($field in @(@('LeftIndent','left'), @('RightIndent','right'), @('FirstLineIndent','firstLine'))) {
                    $measurement = $actual.Format.($field[0])
                    if ([Math]::Abs($measurement - $expectedParagraph.($field[1])) -gt 0.1) { $failures.Add("Indent mismatch at paragraph $($i + 1) in $($case.file): $($field[0]) Word=$measurement, ritr=$($expectedParagraph.($field[1]))") }
                }
                [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($actual)
            }
            Write-Output "Checked $($case.file): $($case.paragraphs.Count) paragraph labels and indents"
        } finally {
            if ($null -ne $document) { $document.Close(0); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($document) }
        }
    }
} finally {
    if ($null -ne $word) { $word.Quit(0); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($word) }
}
if ($failures.Count) { throw ($failures -join "`n") }
Write-Output 'PASS: all labels and indentation match Word'
