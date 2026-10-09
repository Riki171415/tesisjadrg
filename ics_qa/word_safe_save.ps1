$w=[Runtime.InteropServices.Marshal]::GetActiveObject('Word.Application')
foreach($d in $w.Documents){
 if($d.FullName -eq 'C:\Users\PUSBIKES-KEMKES\Downloads\mdc\ICS v9 - Siap Publikasi.docx'){
  $d.Save();Write-Output "Publication document saved: $($d.Saved)"
  if($d.Saved){$d.Close(0);Write-Output 'Saved publication document closed normally.'}
 }
}
