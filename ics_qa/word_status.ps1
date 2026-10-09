$w=[Runtime.InteropServices.Marshal]::GetActiveObject('Word.Application')
$w | Select-Object BackgroundPrintingStatus,BackgroundSavingStatus
Write-Output "Documents $($w.Documents.Count)"
foreach($d in $w.Documents){Write-Output "Document $($d.FullName) saved $($d.Saved)"}
