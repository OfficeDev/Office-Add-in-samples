#!/bin/bash

# Define the base directory to search
BASE_DIR=$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" &> /dev/null && pwd)
echo "Base directory: $BASE_DIR"

# Define the file name or pattern to search for
FILE_PATTERN="package-lock.json"

# Find files and switch to their directories
find "$BASE_DIR" -type f -name "$FILE_PATTERN" | while read -r FILE; do
  DIR=$(dirname "$FILE")  # Extract the directory path
  echo "Switching to directory: $DIR"
  cd "$DIR" || exit       # Change to the directory
  echo "Running npm audit fix (breaking changes will need to be addressed manually)"
  npm audit fix --package-lock-only || true
  cd "$BASE_DIR" || exit  # Return to base directory
done
