# Define the files and directories to include
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

# Create a temporary directory for the ZIP process
$tempDir = New-Item -ItemType Directory -Path (Join-Path $env:TEMP "WordZipTemp") -Force

# Copy the required files and directories to the temporary directory
foreach ($item in $items) {
    Copy-Item -Path $item -Destination $tempDir -Recurse
}

# Define the ZIP and DOCX file paths
$zipFile = Join-Path $PWD "Document.zip"
$docxFile = Join-Path $PWD "Document.docx"

# Create the ZIP file using .NET compression
Add-Type -AssemblyName "System.IO.Compression.FileSystem"
[System.IO.Compression.ZipFile]::CreateFromDirectory($tempDir.FullName, $zipFile)

# Clean up the temporary directory
Remove-Item -Path $tempDir.FullName -Recurse -Force

# Rename the ZIP file to .docx
if (Test-Path $docxFile) { Remove-Item -Path $docxFile -Force }
Rename-Item -Path $zipFile -NewName $docxFile

Write-Host "Document.docx has been created successfully."
