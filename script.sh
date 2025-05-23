#!/bin/bash

EXCEL_FILE="test2.xlsx"
PYTHON_SCRIPT_PATH="main.py"
skipped_or_errored_sheets=()

echo "Starting sheet processing..."
echo "--------------------------------------------------"

# Loop
for i in $(seq 5084 5085)
do
  current_number="$i"
  sheet_identifier_arg="$current_number" 
  output_json_filename="${current_number}.json"

  echo "Running: python3 \"${PYTHON_SCRIPT_PATH}\" \"${EXCEL_FILE}\" \"${sheet_identifier_arg}\" \"${output_json_filename}\""

  python3 "${PYTHON_SCRIPT_PATH}" "${EXCEL_FILE}" "${sheet_identifier_arg}" "output.json"
  
  exit_status=$?
  
  if [ $exit_status -ne 0 ]; then
    echo "WARNING: Sheet ${current_number} - Python script exited with status ${exit_status}."
    skipped_or_errored_sheets+=("${current_number} (Exit Code: ${exit_status})")
  fi
  
  sleep 1 
done

echo "--------------------------------------------------"
echo "All commands executed."
echo ""

if [ ${#skipped_or_errored_sheets[@]} -ne 0 ]; then
  echo "The following sheets were skipped or resulted in an error during processing:"
  for sheet_info in "${skipped_or_errored_sheets[@]}"; do
    echo "  - Sheet ${sheet_info}"
  done
else
  echo "All sheets were processed successfully (or the Python script reported exit code 0 for all)."
fi

echo "--------------------------------------------------"