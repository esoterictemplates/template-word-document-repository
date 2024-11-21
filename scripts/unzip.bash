#!/bin/bash

# Ensure the file exists
if [ ! -f "Document.docx" ]; then
  echo "Error: Document.docx not found!"
  exit 1
fi

# Rename Document.docx to Document.zip
mv "Document.docx" "Document.zip"

# Unzip the renamed file, replacing existing files
unzip -o "Document.zip"

# Clean up (optional: remove the zip file after extraction)
# rm "Document.zip"

echo "Extraction complete."

mv "Document.zip" "Document.docx"
