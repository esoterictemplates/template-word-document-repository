#!/bin/bash

# Find the file matching "Document.*"
file=$(ls Document.* 2>/dev/null)

# Ensure the file exists
if [ -z "$file" ]; then
  echo "Error: No file matching Document.* found!"
  exit 1
fi

# Check if the file is actually a zip archive
if ! file "$file"; then
  echo "Error: $file is not a valid zip archive!"
  exit 1
fi

# Rename the file to Document.zip
mv "$file" "Document.zip"

# Unzip the renamed file, replacing existing files
unzip -o "Document.zip"

# Clean up (optional: remove the zip file after extraction)
# rm "Document.zip"

echo "Extraction complete."

# Rename it back to the original name
mv "Document.zip" "$file"

# Run the Node.js scripts
node scripts/format.js
node scripts/replaceAuthor.js
