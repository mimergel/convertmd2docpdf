# Check if Pandoc is installed
if (-Not (Get-Command pandoc -ErrorAction SilentlyContinue)) {
    Write-Host "Pandoc is not installed or not available in the system's PATH. Please install Pandoc to proceed." -ForegroundColor Red
    exit
}

# Prompt the user for the folder containing markdown files
$markdownFolder = Read-Host "Enter the folder path containing markdown files"

# Resolve and validate path
$markdownFolder = Resolve-Path $markdownFolder | Select-Object -ExpandProperty Path
if (-Not (Test-Path -Path $markdownFolder)) {
    Write-Host "The folder path does not exist. Please provide a valid path." -ForegroundColor Red
    exit
}

# Prompt for optional PDF output
$includePDF = Read-Host "Do you want to also generate PDF files from Word documents? (Yes/No)"
$generatePDF = $includePDF -match '^(?i)y(es)?$'

# Prompt the user for the name of the media or assets folder
$mediaFolderName = Read-Host "Enter the name of the media or assets folder containing images"

# Create output folders
$wordOutput = Join-Path -Path $markdownFolder -ChildPath "word-output"
if (-Not (Test-Path -Path $wordOutput)) {
    New-Item -ItemType Directory -Path $wordOutput | Out-Null
}

$pdfOutput = $null
if ($generatePDF) {
    $pdfOutput = Join-Path -Path $markdownFolder -ChildPath "pdf-output"
    if (-Not (Test-Path -Path $pdfOutput)) {
        New-Item -ItemType Directory -Path $pdfOutput | Out-Null
    }
}

# Create the media folder path
$mediaFolder = Join-Path -Path $markdownFolder -ChildPath $mediaFolderName
if (-Not (Test-Path -Path $mediaFolder)) {
    Write-Host "Could not find the media folder. Please ensure it exists." -ForegroundColor Red
    exit
}

# Convert Markdown to Word (.docx)
Get-ChildItem -Path $markdownFolder -Filter *.md -Recurse | ForEach-Object {
    $mdFile = $_.FullName
    $docxFile = Join-Path -Path $wordOutput -ChildPath ($_.BaseName + ".docx")

    pandoc -o $docxFile -f markdown -t docx $mdFile --extract-media=$mediaFolder --resource-path="$markdownFolder;$mediaFolder"
    Write-Host "Converted $mdFile to $docxFile (with embedded images)"
}

# Optional: Convert Word to PDF
if ($generatePDF) {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0

    Get-ChildItem -Path $wordOutput -Include *.doc, *.docx -Recurse | Where-Object {
        $_.Name -notlike '~$*' -and !$_.PSIsContainer
    } | ForEach-Object {
        $file = $_
        try {
            $doc = $word.Documents.Open($file.FullName, $false, $true)
            $pdfPath = Join-Path $pdfOutput ($file.BaseName + ".pdf")
            $doc.SaveAs($pdfPath, 17)  # 17 = wdFormatPDF
            $doc.Close()
            Write-Host "Converted $($file.Name) to PDF"
        } catch {
            Write-Host "Error processing file: $($file.FullName)" -ForegroundColor Red
            Write-Host $_.Exception.Message
        }
    }

    $word.Quit()
    Start-Process $pdfOutput
}
