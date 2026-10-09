param([switch]$SkipNumbering, [switch]$PagesOnly)
$ErrorActionPreference = 'Stop'
$taskRoot = 'C:\Users\PUSBIKES-KEMKES\Downloads\mdc'
$taskDoc = Join-Path $taskRoot 'ICS v9 - Siap Publikasi.docx'
$taskPdf = Join-Path $taskRoot 'ics_qa\publish-preview.pdf'
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {
    $doc = $word.Documents.Open($taskDoc, $false, $false)
    $changed = 0
    if (-not $SkipNumbering) { foreach ($template in $doc.ListTemplates) {
        if ($template.ListLevels.Item(1).NumberStyle -eq 3) {
            for ($level = 1; $level -le 5; $level++) {
                $ll = $template.ListLevels.Item($level)
                $ll.NumberStyle = @(3,0,4,0,4)[$level-1]
                $ll.NumberFormat = '%' + $level + $(if ($level -le 3) {'.'} else {')'})
                $ll.StartAt = 1
                $ll.ResetOnHigher = $level - 1
                $ll.NumberPosition = ($level - 1) * 21.2598
                $ll.TextPosition = $level * 21.2598
                $ll.TabPosition = $level * 21.2598
                $ll.TrailingCharacter = 0
                $ll.Font.Name = 'Times New Roman'
                $ll.Font.Size = $(if ($level -eq 1) {14} else {12})
            }
            $changed++
        }
    } }
    Write-Output "Five-level list templates: $changed"
    if (-not $PagesOnly) {
    $doc.Fields.Update() | Out-Null
    Write-Output 'Fields updated.'
    foreach ($toc in $doc.TablesOfContents) { $toc.Update() }
    $tocIndex = 0
    foreach ($toc in $doc.TablesOfContents) {
        $tocIndex++
        $toc.Range.Font.Bold = 0
        foreach ($para in $toc.Range.Paragraphs) {
            $styleName = $para.Style.NameLocal
            $levMatch = [regex]::Match($styleName,'(?i)toc\s*([1-9])')
            if (-not $levMatch.Success) { continue }
            $lev = [int]$levMatch.Groups[1].Value - 1
            $para.Range.Font.Name = 'Times New Roman'
            $para.Range.Font.Size = $(if ($tocIndex -eq 1) {12} else {10})
            $para.Format.SpaceBefore = 0
            $para.Format.SpaceAfter = 2
            $para.Format.LineSpacingRule = 0
            $para.Format.Alignment = 0
            $para.Format.KeepWithNext = 0
            $para.Format.KeepTogether = 0
            if ($tocIndex -eq 1 -and $para.Range.Text -match '^\s*([A-Z]\.|\d+[.)]|[a-z]\.)\s') {
                $para.Format.LeftIndent = ($lev + 1) * 21.2598
                $para.Format.FirstLineIndent = -21.2598
            } else {
                $para.Format.LeftIndent = 0
                $para.Format.FirstLineIndent = 0
            }
        }
    }
    Write-Output 'TOC indentation normalized.'
    }
    $doc.Repaginate()
    foreach ($toc in $doc.TablesOfContents) { $toc.UpdatePageNumbers() }
    $titleRange = $doc.Content.Duplicate
    $titleRange.Find.Text = 'DAFTAR ISI'
    $titleRange.Find.Wrap = 0
    if ($titleRange.Find.Execute()) {
        $titleRange.Font.Name = 'Times New Roman'
        $titleRange.Font.Size = 14
        $titleRange.Font.Bold = -1
    }
    $doc.Save()
    Write-Output "Pages: $($doc.ComputeStatistics(2))"
    $doc.ExportAsFixedFormat($taskPdf,17)
    $doc.Close(0)
    Write-Output 'Saved Word document and verification PDF.'
} finally {
    $word.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
}
