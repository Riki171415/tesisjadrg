$w=[Runtime.InteropServices.Marshal]::GetActiveObject('Word.Application')
if($w.Documents.Count -eq 0){$w.Quit();Write-Output 'Empty automation Word instance closed normally.'}else{Write-Output 'Word has an open document; left running.'}
