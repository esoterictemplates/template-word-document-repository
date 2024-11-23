# Prompt the user for the desired file extension
$fileExtension = Read-Host "Enter the desired file extension (e.g., docx, pptx, xlsx)"

# Validate the input
if (-not $fileExtension -or $fileExtension -match '[^\w]') {
    Write-Host "Error: Invalid file extension provided. Exiting." -ForegroundColor Red
    exit 1
}

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

# Path to the 7-Zip executable (quoted)
$sevenZipPath = '"C:\Program Files\7-Zip\7z.exe"'

# Verify 7-Zip is installed
if (-not (Test-Path $sevenZipPath -replace '"', '')) {
    Write-Host "Error: 7-Zip not found at $sevenZipPath. Please install it and try again." -ForegroundColor Red
    exit 1
}

# Create the ZIP file
$zipFile = Join-Path $PWD "Document.zip"
$outputFile = Join-Path $PWD "Document.$fileExtension"

# Build the 7-Zip command
$command = @("$sevenZipPath", "a", "-tzip", "`"$zipFile`"")

# Add the items to the command (each item quoted)
foreach ($item in $items) {
    $command += "`"$item`""
}

# Run the 7-Zip command
try {
    # Join the arguments into a single string for Start-Process
    $arguments = $command[1..($command.Count - 1)] -join " "
    Start-Process -FilePath $sevenZipPath -ArgumentList $arguments -Wait -NoNewWindow
    Write-Host "Created $zipFile using 7-Zip."
} catch {
    Write-Host "Error: Failed to create $zipFile. Exiting." -ForegroundColor Red
    exit 1
}

# Rename the ZIP file to the specified file extension
if (Test-Path $outputFile) { Remove-Item -Path $outputFile -Force }
Rename-Item -Path $zipFile -NewName $outputFile

Write-Host "Document.$fileExtension has been created successfully."
