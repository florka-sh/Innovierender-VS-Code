@echo off
REM Batch wrapper script to set Poppler in PATH before running InvoiceExtractor

setlocal enabledelayedexpansion

REM Get the directory where this batch file is located
set "scriptDir=%~dp0"

REM Add the bundled Poppler to the PATH
set "popplerPath=%scriptDir%_internal\poppler\Library\bin"
set "PATH=%popplerPath%;!PATH!"

REM Add the bundled Tesseract to the PATH
set "tesseractPath=%scriptDir%_internal\Tesseract-OCR"
set "PATH=%tesseractPath%;!PATH!"

REM Also set environment variables that some PDF libraries use
set "POPPLER_HOME=%scriptDir%_internal\poppler"
set "TESSDATA_PREFIX=%scriptDir%_internal\Tesseract-OCR\tessdata"

echo Poppler PATH set to: %popplerPath%
echo Tesseract PATH set to: %tesseractPath%
echo Starting InvoiceExtractor with Poppler support...

REM Run the executable
"%scriptDir%InvoiceExtractor.exe"

endlocal
