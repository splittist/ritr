param(
    [Parameter(Mandatory = $true)][string]$Path,
    [string]$ExpectedText = 'Written confirmation'
)

# Use a new hidden Word instance, read-only; never attach to the user's session.
$ErrorActionPreference = 'Stop'
$word = $null
$document = $null
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $word.AutomationSecurity = 3
    $document = $word.Documents.Open((Resolve-Path -LiteralPath $Path).Path, $false, $true, $false)
    $table = $document.Tables.Item(1)
    if (-not $table.Cell(1, 1).Range.Text.Contains('RESPONSIBILITIES AND TERMS')) { throw 'Missing spanning header' }
    if (-not $table.Cell(2, 1).Range.Text.Contains('Delivery')) { throw 'Missing vertically merged cell' }
    if (-not $table.Cell(3, 3).Range.Text.Contains($ExpectedText)) { throw 'Unexpected confirmation cell text' }
    if (-not $table.Cell(4, 3).Range.Text.Contains('Currency')) { throw 'Missing nested table' }
    [pscustomobject]@{ File = $Path; Rows = $table.Rows.Count; ReadOnly = $document.ReadOnly; TableChecks = 'passed' }
} finally {
    if ($null -ne $document) { $document.Close(0); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($document) }
    if ($null -ne $word) { $word.Quit(0); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($word) }
}
