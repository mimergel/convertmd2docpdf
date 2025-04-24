# Convert MD files to Word and PDF documents
Converts MD files to word including images in a to be specified subfolder and optionally creates PDF from the word documents

## Prerequisites

1. **PowerShell 5.1 or later**.
2. [**Pandoc**](https://pandoc.org/installing.html) (must be installed and accessible in your system's PATH).
3. **Microsoft Word**.

## Installation

1. **Clone the repository** (or download the ZIP) to your local machine:
   ```powershell
   git clone https://github.com/mimergel/convertmd2docpdf.git
   ```
2. **Navigate** to the folder where you cloned or unzipped the repository.


## Running the script

1. **Open PowerShell** in the repository folder.
2. **Run** the script 
   ```powershell
   .\convertmd2docpdf.ps1
   ```
3. **Follow** the on-screen prompts to specify:
   ```
   Enter the folder path containing markdown files: [folder to MD files, e.g. md_files]
   Do you want to also generate PDF files from Word documents? (Yes/No): Yes
   Enter the name of the media or assets folder containing images: [subfolder name within the MD files folder, e.g. assets]
   Converted filename1.md to C:\Users\mimergel\OneDrive - filename1.docx (with embedded images)
   Converted filename2.md to C:\Users\mimergel\OneDrive - filename2.docx (with embedded images)
   Converted filename1.docx to PDF
   Converted filename2.docx to PDF
   ```
## Troubleshooting

- If `pandoc` is not recognized, ensure Pandoc is installed and the install location is in your PATH.
- If Microsoft Word is not installed, you cannot generate PDFs; only Word documents will be created.

## License
This script is licensed under the [MIT License](LICENSE).  
You are free to use, modify, and distribute it without restriction.
