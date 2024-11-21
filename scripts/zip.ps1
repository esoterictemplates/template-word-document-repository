# Define the files and directories to include
$items = @("_rels", "docProps", "word", "[Content_Types].xml")

# Check if all required items exist
$missing = $false
foreach ($item in $items) {
    Write-Host "Checking $item..."
    if (-not (Test-Path $item)) {
        Write-Host "Error: $item not found!" -ForegroundColor Red
        $missing = $true
    }
}

if ($missing) {
    Write-Host "One or more required files/directories are missing. Exiting." -ForegroundColor Red
    exit 1
}

# Create the zip file
$zipFile = "Document.zip"
try {
    Compress-Archive -Path $items -DestinationPath $zipFile -Force
    Write-Host "Created $zipFile."
} catch {
    Write-Host "Error: Failed to create $zipFile. Exiting." -ForegroundColor Red
    exit 1
}

# Rename the zip file to Document.docx
$docxFile = "Document.docx"
try {
    Rename-Item -Path $zipFile -NewName $docxFile -Force
    Write-Host "$docxFile has been created."
} catch {
    Write-Host "Error: Failed to rename $zipFile to $docxFile. Exiting." -ForegroundColor Red
    exit 1
}
