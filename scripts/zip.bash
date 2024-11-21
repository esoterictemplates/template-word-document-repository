#!/bin/bash

# Check if the required files and directories exist
missing=false
for item in "_rels" "docProps" "word" "[Content_Types].xml"; do
  if [ ! -e "$item" ]; then
    echo "Error: $item not found!"
    missing=true
  fi
done

if [ "$missing" = true ]; then
  echo "One or more required files/directories are missing. Exiting."
  exit 1
fi

# Create the zip file
zip -r "Document.zip" "_rels" "docProps" "word" "[Content_Types].xml"

# Rename the zip file to Document.docx
mv "Document.zip" "Document.docx"

echo "Document.docx has been created."
