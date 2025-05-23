import argparse
import os
import json
from aexcel import excel_to_json
from bnulls import clean_json_data
from cformat import format_dynamically


def run_scripts(excel_file, output_file, sheet_name):
    # Run aexcel.py
    temp_json_file = 'temp.json'
    excel_to_json(excel_file, temp_json_file, sheet_name)

    # Run bnulls.py
    with open(temp_json_file, 'r', encoding='utf-8') as f:
        json_data = json.load(f)

    cleaned_data = clean_json_data(json_data)
    formatted_data = format_dynamically(cleaned_data)
    
    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(formatted_data, f, indent=4, ensure_ascii=False)

def main():
    parser = argparse.ArgumentParser(description="Run all scripts and output final JSON.")
    parser.add_argument("excel_file", help="Path to the Excel file.")
    parser.add_argument("sheet_name", help="Name of the sheet to convert.")
    parser.add_argument("output_file", help="Path to save the final JSON output.")
    args = parser.parse_args()

    run_scripts(args.excel_file, args.output_file, args.sheet_name)


if __name__ == "__main__":
    main()