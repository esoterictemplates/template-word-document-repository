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
    "LICENSE", "package.json", "README.md", "package-lock.json"
)

# Convert $PWD to a string
$currentDir = $PWD.ProviderPath

# Get all items in the current directory excluding the specified ones
$items = Get-ChildItem -Path $currentDir -Recurse | Where-Object {
    $relativePath = $_.FullName.Substring($currentDir.Length).TrimStart("\", "/")
    
    # Debugging: Show the relative path of each file
    Write-Host "Checking file: $relativePath"

    for ($i = 0; $i -lt $excludedItems.Length; $i++) {
        if ($relativePath -like "*$($excludedItems[$i])*") {
            return $false
        }
    }

    # Exclude items if they match any of the exclusions
    -not ($excludedItems -contains $relativePath)
} | ForEach-Object {
    # Only return files (not directories)
    if ($_ -is [System.IO.FileInfo]) {
        $_.FullName
    }
}

# Check if there are items to process
if (-not $items) {
    Write-Host "Error: No items found to include in the archive. Exiting." -ForegroundColor Red
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
$zipFile = Join-Path $currentDir "Document.zip"
$outputFile = Join-Path $currentDir "Document.$fileExtension"

# Ensure no pre-existing ZIP file blocks creation
if (Test-Path $zipFile) { Remove-Item -Path $zipFile -Force }

# Create a temporary file list for 7-Zip to avoid duplicate file names
$tempFileList = Join-Path $currentDir "filelist.txt"

# Write relative file paths to the temporary list
$items | ForEach-Object {
    $relativePath = $_.Substring($currentDir.Length).TrimStart("\", "/")
    "`"$relativePath`""  # Add quotes around the relative paths
} | Set-Content -Path $tempFileList -Encoding UTF8

# Build and run the 7-Zip command
try {
    $arguments = @("a", "-tzip", "`"$zipFile`"", "-spf2", "@`"$tempFileList`"")
    Start-Process -FilePath $sevenZipPath -ArgumentList $arguments -Wait -NoNewWindow
    Write-Host "Created $zipFile using 7-Zip."
} catch {
    Write-Host "Error: Failed to create $zipFile. Exiting." -ForegroundColor Red
    exit 1
}

# Clean up the temporary file list
if (Test-Path $tempFileList) { Remove-Item -Path $tempFileList -Force }

# Rename the ZIP file to the specified file extension
if (Test-Path $outputFile) { Remove-Item -Path $outputFile -Force }
Rename-Item -Path $zipFile -NewName $outputFile

Write-Host "Document.$fileExtension has been created successfully."
