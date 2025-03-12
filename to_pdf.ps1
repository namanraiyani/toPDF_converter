$folderPath = Split-Path -Parent $MyInvocation.MyCommand.Path

$pptxFiles = Get-ChildItem -Path $folderPath -Filter *.pptx
$docxFiles = Get-ChildItem -Path $folderPath -Filter *.docx

$powerPoint = New-Object -ComObject PowerPoint.Application
$powerPoint.Visible = [Microsoft.Office.Core.MsoTriState]::msoFalse

$word = New-Object -ComObject Word.Application
$word.Visible = $false

foreach ($pptxFile in $pptxFiles) {
    $pdfFile = [System.IO.Path]::ChangeExtension($pptxFile.FullName, ".pdf")
    try {
        $presentation = $powerPoint.Presentations.Open($pptxFile.FullName)
        $presentation.SaveAs($pdfFile, 32)
        $presentation.Close()
        Write-Host "Converted $($pptxFile.Name) to PDF successfully!"
    } catch {
        Write-Host "Failed to convert $($pptxFile.Name) to PDF. Error: $_"
    }
}

foreach ($docxFile in $docxFiles) {
    $pdfFile = [System.IO.Path]::ChangeExtension($docxFile.FullName, ".pdf")
    try {
        $document = $word.Documents.Open($docxFile.FullName)
        $document.SaveAs([ref]$pdfFile, [ref]17)
        $document.Close()
        Write-Host "Converted $($docxFile.Name) to PDF successfully!"
    } catch {
        Write-Host "Failed to convert $($docxFile.Name) to PDF. Error: $_"
    }
}

$powerPoint.Quit()
$word.Quit()
