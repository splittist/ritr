param(
    [Parameter(Mandatory = $true)][string[]]$Paths,
    [string]$ExpectedText = ''
)

# Interoperability check using an isolated, hidden Word instance. Never attaches
# to the user's running Word session and never writes the input documents.
$ErrorActionPreference = 'Stop'
$word = $null
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $word.AutomationSecurity = 3
    foreach ($inputPath in $Paths) {
        $resolved = (Resolve-Path -LiteralPath $inputPath).Path
        $document = $null
        try {
            $document = $word.Documents.Open($resolved, $false, $true, $false)
            $text = $document.Content.Text
            if ($ExpectedText -and -not $text.Contains($ExpectedText)) {
                throw "Word did not find expected text in $resolved"
            }
            [pscustomobject]@{ File = $resolved; Paragraphs = $document.Paragraphs.Count; ReadOnly = $document.ReadOnly; ExpectedTextFound = (-not $ExpectedText -or $text.Contains($ExpectedText)) }
        } finally {
            if ($null -ne $document) {
                $document.Close(0)
                [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($document)
            }
        }
    }
} finally {
    if ($null -ne $word) {
        $word.Quit(0)
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($word)
    }
}
