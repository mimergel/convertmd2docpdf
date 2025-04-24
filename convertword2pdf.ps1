# Prompt the user for the folder containing Word files
$sourceFolder = Read-Host "Enter the folder path containing Word files"

# Resolve full path and validate
$sourceFolder = Resolve-Path $sourceFolder | Select-Object -ExpandProperty Path
if (-Not (Test-Path -Path $sourceFolder)) {
    Write-Host "The folder path does not exist. Please provide a valid path." -ForegroundColor Red
    exit
}

# Create output folder for PDFs
$targetFolder = Join-Path -Path $sourceFolder -ChildPath "pdf-output"
if (-Not (Test-Path -Path $targetFolder)) {
    New-Item -ItemType Directory -Path $targetFolder | Out-Null
}

# Start Word COM object
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0  # Suppress dialogs

# Loop through valid .doc/.docx files (exclude temp files)
Get-ChildItem -Path $sourceFolder -Include *.doc, *.docx -Recurse | Where-Object {
    $_.Name -notlike '~$*' -and !$_.PSIsContainer
} | ForEach-Object {
    $file = $_
    try {
        $doc = $word.Documents.Open($file.FullName, $false, $true)
        $pdfPath = Join-Path $targetFolder ($file.BaseName + ".pdf")
        $doc.SaveAs($pdfPath, 17)  # 17 = wdFormatPDF
        $doc.Close(0)
    } catch {
        Write-Host "Error processing file: $($file.FullName)" -ForegroundColor Red
        Write-Host $_.Exception.Message
    }
}

# Quit Word
$word.Quit()

# Open output folder
Start-Process $targetFolder
