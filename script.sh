#!/bin/bash

# The fixed Excel filename
EXCEL_FILE="test2.xlsx"

# The base path for your Python script (if not in PATH or current dir)
# PYTHON_SCRIPT_PATH="./combined_script.py" # Example if it's in the current directory
PYTHON_SCRIPT_PATH="main.py" # Assuming it's in PATH or current directory

# Loop from 5000 to 5099
for i in $(seq 5000 5099)
do
  # The number to be used for sheet identification and JSON filename
  current_number="$i"
  output_json_filename="${current_number}.json"

  # Echo the command that will be run (optional, useful for logging/debugging)
  echo "Running: python3 \"${PYTHON_SCRIPT_PATH}\" \"${EXCEL_FILE}\" \"${current_number}\" \"${output_json_filename}\""

  # Execute the Python script
  python3 "${PYTHON_SCRIPT_PATH}" "${EXCEL_FILE}" "${current_number}" "${output_json_filename}"
  
  # Optional: Add a small delay if needed, e.g., sleep 1 (for 1 second)
done

echo "All commands executed."