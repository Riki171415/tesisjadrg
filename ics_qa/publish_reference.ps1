$ErrorActionPreference = 'Stop'
$taskRoot = 'C:\Users\PUSBIKES-KEMKES\Downloads\mdc'
$taskDoc = Join-Path $taskRoot 'ICS v9 - Siap Publikasi.docx'
$patchPdf = Join-Path $taskRoot 'ics_qa\reference-page.pdf'
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {
    $doc = $word.Documents.Open($taskDoc,$false,$false)
    $refRange = $doc.Content.Duplicate
    $refRange.Find.Text = 'Pasien dalam contoh 8'
    $refRange.Find.Wrap = 0
    if (-not $refRange.Find.Execute()) { throw 'Expected example reference was not found.' }
    $refRange.Text = 'Pasien dalam contoh 5'
    $pageNumber = [int]$refRange.Information(3)
    $doc.Save()
    $doc.ExportAsFixedFormat($patchPdf,17,$false,0,3,$pageNumber,$pageNumber)
    $pageCount = [int]$doc.ComputeStatistics(2)
    @{ page=$pageNumber; pages=$pageCount } | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $taskRoot 'ics_qa\reference-page.json')
    $doc.Close(0)
    Write-Output "Reference corrected; rendered page $pageNumber; total pages $pageCount."
} finally {
    $word.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
}
