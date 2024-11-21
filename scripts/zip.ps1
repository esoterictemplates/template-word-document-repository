# Define the paths to the files and directories
$items = @("_rels", "docProps", "word", "[Content_Types].xml")

# Check if all required items exist
$missing = $false
foreach ($item in $items) {
    Write-Host "Checking $item..."
    # Escape special characters in the item name
    $escapedItem = $item -replace '\[', '`[' -replace '\]', '`]'
    if (-not (Test-Path $escapedItem)) {
        Write-Host "Error: $item not found!" -ForegroundColor Red
        $missing = $true
    }
}

if ($missing) {
    Write-Host "One or more required files/directories are missing. Exiting." -ForegroundColor Red
    exit 1
}

# Path to the 7-Zip executable
$sevenZipPath = "C:\Program Files\7-Zip\7z.exe"

# Verify 7-Zip is installed
if (-not (Test-Path $sevenZipPath)) {
    Write-Host "Error: 7-Zip not found at $sevenZipPath. Please install it and try again." -ForegroundColor Red
    exit 1
}

# Create the ZIP file
$zipFile = Join-Path $PWD "Document.zip"
$docxFile = Join-Path $PWD "Document.docx"

# Build the 7-Zip command
$command = @("$sevenZipPath", "a", "-tzip", $zipFile)
$command += $items

# Run the 7-Zip command
try {
    Start-Process -FilePath $sevenZipPath -ArgumentList $command -Wait -NoNewWindow
    Write-Host "Created $zipFile using 7-Zip."
} catch {
    Write-Host "Error: Failed to create $zipFile. Exiting." -ForegroundColor Red
    exit 1
}

# Rename the ZIP file to .docx
if (Test-Path $docxFile) { Remove-Item -Path $docxFile -Force }
Rename-Item -Path $zipFile -NewName $docxFile

Write-Host "Document.docx has been created successfully and should now open in Word."
