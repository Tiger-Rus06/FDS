# Refreshes page numbers in the report with Microsoft Word and saves a PDF copy.
# Usage: powershell -ExecutionPolicy Bypass -File export_pdf.ps1
$docx = [string](Join-Path $PSScriptRoot "Data_Science_Case_Studies_Report.docx")
$pdf = [string](Join-Path $PSScriptRoot "Data_Science_Case_Studies_Report.pdf")
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {
    $doc = $word.Documents.Open($docx, $false, $false, $false)
    $doc.Repaginate()
    $doc.Fields.Update() | Out-Null
    $doc.ExportAsFixedFormat($pdf, 17)
    $doc.Save()
    $doc.Close(0)
    Write-Output "Saved $pdf"
} finally {
    $word.Quit()
}
