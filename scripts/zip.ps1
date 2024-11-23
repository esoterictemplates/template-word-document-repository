# Prompt the user for the desired file extension
$fileExtension = Read-Host "Enter the desired file extension (e.g., docx, pptx, xlsx)"

# Validate the input
if (-not $fileExtension -or $fileExtension -match '[^\w]') {
    Write-Host "Error: Invalid file extension provided. Exiting." -ForegroundColor Red
    exit 1
}

# Define the exclusions
$excludedItems = @(
    ".git", ".vscode", "assets", "node_modules", "scripts",
    ".gitignore", "CHANGELOG.md", "CODE_OF_CONDUCT.md", 
    "LICENSE", "package.json", "README.md"
)

# Get all items in the current directory excluding the specified ones
$items = Get-ChildItem -Path $PWD -Recurse -File | Where-Object {
    $relativePath = $_.FullName -replace "$PWD[\\\/]?", ""
    -not ($relativePath -in $excludedItems)
} | ForEach-Object { $_.FullName }

# Check if there are items to process
if (-not $items) {
    Write-Host "Error: No items found to include in the archive. Exiting." -ForegroundColor Red
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
    $command += "`"$item`""  # Add escaped paths to the command
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
