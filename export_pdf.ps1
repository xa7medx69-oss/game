$ErrorActionPreference = 'Stop'
$docx = 'C:\Users\Ahmed\Documents\ChatGPT\games\physical-social-games-handbook\Physical-Social-Games-Business-Handbook.docx'
$pdf = 'C:\Users\Ahmed\Documents\ChatGPT\games\physical-social-games-handbook\Physical-Social-Games-Business-Handbook.pdf'
$word = New-Object -ComObject Word.Application
try {
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $doc = $word.Documents.Open($docx, $false, $true)
    $pages = $doc.ComputeStatistics(2)
    $doc.ExportAsFixedFormat($pdf, 17)
    $doc.Close($false)
    Write-Output "WORD_EXPORT_OK pages=$pages pdf=$pdf"
}
finally {
    $word.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
}
