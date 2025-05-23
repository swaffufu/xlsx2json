import pandas as pd
import argparse


def excel_to_json(excel_file, output_file=None, sheet_name='Sheet'):
    try:
        excel_data_fragment = pd.read_excel(excel_file, sheet_name=sheet_name)
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