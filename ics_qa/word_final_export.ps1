$w=New-Object -ComObject Word.Application
$w.Visible=$false
$w.DisplayAlerts=0
try{
 $d=$w.Documents.Open('C:\Users\PUSBIKES-KEMKES\Downloads\mdc\ICS v9 - Siap Publikasi.docx',$false,$false)
 $d.Repaginate()
 foreach($toc in $d.TablesOfContents){$toc.UpdatePageNumbers()}
 $d.Save()
 Write-Output "Saved; pages $($d.ComputeStatistics(2))"
 $d.ExportAsFixedFormat('C:\Users\PUSBIKES-KEMKES\Downloads\mdc\ics_qa\final-preview.pdf',17,$false,0,0,1,1,0,$true,$true,0,$false,$true,$false)
 Write-Output 'PDF export completed.'
 $d.Close(0)
}finally{$w.Quit()}
