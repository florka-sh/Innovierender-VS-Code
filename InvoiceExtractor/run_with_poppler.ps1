# PowerShell wrapper script to set Poppler in PATH before running InvoiceExtractor

# Get the directory where this script is located
$scriptDir = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition

# Add the bundled Poppler to the PATH
$popplerPath = Join-Path $scriptDir "_internal\poppler\Library\bin"
$env:PATH = "$popplerPath;$env:PATH"

# Add the bundled Tesseract to the PATH
$tesseractPath = Join-Path $scriptDir "_internal\Tesseract-OCR"
$env:PATH = "$tesseractPath;$env:PATH"

# Also set environment variables that some PDF libraries use
$env:POPPLER_HOME = Join-Path $scriptDir "_internal\poppler"
$env:TESSDATA_PREFIX = Join-Path $scriptDir "_internal\Tesseract-OCR\tessdata"

Write-Host "Poppler PATH set to: $popplerPath"
Write-Host "Tesseract PATH set to: $tesseractPath"
Write-Host "Starting InvoiceExtractor with Poppler support..."

# Run the executable
& "$scriptDir\InvoiceExtractor.exe"
