$ErrorActionPreference = 'Stop'
$projectRoot = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
$docx = Join-Path $projectRoot 'manuals\compiled\Physical-Social-Games-Business-Handbook.docx'
$pdf = Join-Path $projectRoot 'manuals\compiled\Physical-Social-Games-Business-Handbook.pdf'
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
