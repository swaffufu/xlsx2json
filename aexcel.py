import pandas as pd
import argparse


def excel_to_json(excel_file, output_file=None, sheet_name='Sheet'):
    try:
        xls = pd.ExcelFile(excel_file)
        all_sheet_names = xls.sheet_names
        
        target_sheet_name = None
        for s_name in all_sheet_names:
            if s_name.startswith(str(sheet_name)):
                target_sheet_name = s_name
                break
        
        if target_sheet_name is None:
            if sheet_name in all_sheet_names:
                target_sheet_name = sheet_name
            else:
                print(f"Error: Worksheet starting with or named '{sheet_name}' not found in '{excel_file}'. Available sheets: {all_sheet_names}")
                return

        excel_data_fragment = pd.read_excel(xls, sheet_name=target_sheet_name)
        json_str = excel_data_fragment.to_json()

        if output_file:
            with open(output_file, 'w', encoding='utf-8') as f:
                f.write(json_str)
            print(f'JSON successfully saved to {output_file}')
        else:
            print('Excel Sheet to JSON:\n', json_str)
    except FileNotFoundError:
        print(f"Error: File '{excel_file}' not found.")
    except Exception as e:
        print(f"An error occurred: {e}")


def main():
    parser = argparse.ArgumentParser(description="Convert Excel file to JSON.")
    parser.add_argument("excel_file", help="Path to the Excel file.")
    parser.add_argument("-o", "--output_file", help="Path to save the JSON output.")
    parser.add_argument("-s", "--sheet_name", help="Name of the sheet to convert.")
    args = parser.parse_args()
    
    excel_to_json(args.excel_file, args.output_file, args.sheet_name)


if __name__ == "__main__":
    main()