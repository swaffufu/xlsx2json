import pandas as pd
import json
import argparse
import sys
from datetime import datetime, timedelta
import re

KOPERASI_KEY = "              KOPERASI PERMODALAN DAN PERUSAHAAN MELAYU NEGERI SEMBILAN BERHAD"

MEMBER_LABELS_MAP = {
    "NO. ANGGOTA": "NO. ANGGOTA", "GELARAN": "GELARAN", "NAMA": "NAMA",
    "NO. K/P": "NO. K/P", "TARIKH LAHIR": "TARIKH LAHIR", "ALAMAT TETAP": "ALAMAT TETAP",
    "ALAMAT SURAT MENYURAT": "ALAMAT SURAT MENYURAT", 
    "NO. TELEFON ANGGOTA": "NO. TELEFON ANGGOTA", "PERKERJAAN": "PERKERJAAN",
    "PENAMA / K.P": "PENAMA_KP_RAW",
    "NO. TELEFON PENAMA ": "NO_TELEFON_PENAMA_RAW",
    "TARIKH MASUK": "TARIKH MASUK", "TARIKH LULUS ALK": "TARIKH LULUS ALK"
}

MEMBER_LABELS_EXPECTED_ROW_INDICES = {
    "NO. ANGGOTA": "5", "GELARAN": "6", "NAMA": "7", "NO. K/P": "8",
    "TARIKH LAHIR": "9", "ALAMAT TETAP": "10", "ALAMAT SURAT MENYURAT": "11",
    "NO. TELEFON ANGGOTA": "12", "PERKERJAAN": "13", "PENAMA / K.P": "14",
    "NO. TELEFON PENAMA ": "15", "TARIKH MASUK": "16", "TARIKH LULUS ALK": "17"
}

TRANSACTION_HEADERS_MAP = {
    "TARIKH": "TARIKH", "PERKARA": "PERKARA", "NO.RESIT": "NO.RESIT", "TAHUN": "TAHUN",
    "WANG MASUK": "WANG MASUK", "WANG KELUAR": "WANG KELUAR",
    "BAKI SYER": "BAKI SYER", "BAKI BONUS": "BAKI BONUS", "BAKI SEMUA": "BAKI SEMUA",
    "CATATAN": "CATATAN"
}

MULTI_PART_TX_HEADERS = {
    "WANG": {"MASUK": "WANG MASUK"},
    "BAKI": {"SYER": "BAKI SYER", "BONUS": "BAKI BONUS", "SEMUA": "BAKI SEMUA"}
}

TX_HEADER_TYPICAL_ROW_INDICES = ("19", "20")


def convert_excel_timestamp(value, date_format='%d-%b-%y'):
    if value is None:
        return None
    
    if isinstance(value, str):
        # Attempt to parse common string date formats
        try:
            # Common date formats, including ISO formats that pandas might produce
            common_formats = [
                '%d-%m-%y', '%d/%m/%y', '%d-%b-%y', 
                '%Y-%m-%d', '%m/%d/%Y', '%d-%m-%Y', '%d/%m/%Y',
                '%Y-%m-%dT%H:%M:%S.%fZ', '%Y-%m-%dT%H:%M:%SZ', # ISO formats
                '%d %b %Y', '%b %d, %Y'
            ]
            dt_object = None
            for fmt in common_formats:
                try:
                    dt_object = datetime.strptime(value, fmt)
                    break 
                except ValueError:
                    continue
            
            if dt_object:
                return dt_object.strftime(date_format).upper()
            # If no format matched, return the string, possibly uppercased
            return value.upper() 
        except Exception:
            return str(value).upper() # Fallback for unexpected string parsing errors

    if isinstance(value, (int, float)):
        try:
            # Case 1: Excel serial date (numbers typically in a specific range like 10000-70000)
            # Ensure it's an integer-like value for Excel days.
            if 10000 <= value <= 70000 and value == int(value):
                dt_object = datetime(1899, 12, 30) + timedelta(days=int(value))
                return dt_object.strftime(date_format).upper()

            # Case 2: Millisecond timestamp (can be positive or negative)
            # Python's datetime objects support years from 1 to 9999.
            # Corresponding seconds from epoch range:
            min_supported_seconds = -62135596800  # datetime(1,1,1).timestamp()
            max_supported_seconds = 253402300799 # datetime(9999,12,31,23,59,59).timestamp()
            
            seconds_candidate = value / 1000.0

            # Heuristic: Check if abs(value) is large (e.g., > 1 day in ms)
            # and if the seconds_candidate is within the range supported by datetime objects.
            if abs(value) > 86400000 and min_supported_seconds <= seconds_candidate <= max_supported_seconds:
                try:
                    # datetime.fromtimestamp can handle negative values on POSIX systems
                    dt_object = datetime.fromtimestamp(seconds_candidate)
                except OSError:
                    # Fallback for systems where fromtimestamp fails for negative values (e.g., Windows pre-epoch)
                    # This method works for both positive and negative seconds relative to 1970-01-01
                    dt_object = datetime(1970, 1, 1) + timedelta(seconds=seconds_candidate)
                return dt_object.strftime(date_format).upper()
            
            # Fallback for other numbers not matching the above criteria
            return str(int(value)) if isinstance(value, float) and value == int(value) else str(value)

        except (ValueError, TypeError, OSError, OverflowError): # Catch errors from datetime/timedelta operations
            return str(value) # Fallback to string if any conversion error occurs
            
    # Default fallback for any other types or unhandled cases
    return str(value)


def parse_nominee_string(nominee_str):
    if not nominee_str or not isinstance(nominee_str, str):
        return {"NAME": None, "RELATIONSHIP": None, "IC": None}
    match = re.match(r"^(.*?)\s*\((.*?)/(.*?)\)$", nominee_str.strip())
    if match:
        name, relationship, ic = match.groups()
        return {"NAME": name.strip(),"RELATIONSHIP": relationship.strip(),"IC": ic.strip()}
    return {"NAME": nominee_str.strip(), "RELATIONSHIP": None, "IC": None}


def find_member_data(raw_data, koperasi_col_obj):
    member_details = {}
    raw_nominee_details = {}
    
    member_value_col_key_found = None
    no_anggota_label_row = MEMBER_LABELS_EXPECTED_ROW_INDICES.get("NO. ANGGOTA")
    
    if no_anggota_label_row and koperasi_col_obj.get(no_anggota_label_row) == "NO. ANGGOTA":
        row_indices_to_check_for_value = [
            no_anggota_label_row,
            str(int(no_anggota_label_row) + 1),
            str(int(no_anggota_label_row) - 1) 
        ]
        preferred_value_cols = ["Unnamed: 3", "Unnamed: 4"] 
        other_potential_value_cols = [k for k in raw_data.keys() if k.startswith("Unnamed:") and k not in preferred_value_cols]
        
        for r_idx_val_check in row_indices_to_check_for_value:
            for col_key_cand in preferred_value_cols + other_potential_value_cols:
                col_obj_cand = raw_data.get(col_key_cand, {})
                val = col_obj_cand.get(r_idx_val_check)
                if isinstance(val, int) and 1000 <= val <= 99999:
                    member_value_col_key_found = col_key_cand
                    member_details[MEMBER_LABELS_MAP["NO. ANGGOTA"]] = val
                    no_anggota_value_row_idx = r_idx_val_check 
                    break
            if member_value_col_key_found:
                break
    
    if not member_value_col_key_found:
        print("[WARNING] Could not dynamically identify the primary Member Value Column. Member details might be incomplete.")
        if raw_data.get("Unnamed: 4", {}).get(MEMBER_LABELS_EXPECTED_ROW_INDICES.get("NAMA")):
            member_value_col_key_found = "Unnamed: 4"
        elif raw_data.get("Unnamed: 3", {}).get(MEMBER_LABELS_EXPECTED_ROW_INDICES.get("NAMA")):
             member_value_col_key_found = "Unnamed: 3"


    member_value_col_data = raw_data.get(member_value_col_key_found, {}) if member_value_col_key_found else {}


    for label_text, output_key in MEMBER_LABELS_MAP.items():
        if output_key in member_details:
            continue

        expected_label_row_idx = MEMBER_LABELS_EXPECTED_ROW_INDICES.get(label_text)
        if not expected_label_row_idx or koperasi_col_obj.get(expected_label_row_idx) != label_text:
            continue

        value = member_value_col_data.get(expected_label_row_idx)

        if label_text == "GELARAN" and value is None and member_value_col_key_found == "Unnamed: 4":
            if no_anggota_value_row_idx == str(int(MEMBER_LABELS_EXPECTED_ROW_INDICES["NO. ANGGOTA"]) + 1) :
                value = member_value_col_data.get(str(int(expected_label_row_idx) - 1))


        if output_key == "PENAMA_KP_RAW":
            raw_nominee_details["PENAMA_KP_RAW"] = value
        elif output_key == "NO_TELEFON_PENAMA_RAW":
            raw_nominee_details["NO_TELEFON_PENAMA_RAW"] = value
        elif "TARIKH" in output_key:
            member_details[output_key] = convert_excel_timestamp(value, '%d-%b-%y')
        else:
            member_details[output_key] = value
            
    return member_details, raw_nominee_details


def find_transaction_columns_and_parse(raw_data, koperasi_col_obj):
    located_tx_cols = {}
    
    for col_key, col_data in raw_data.items():
        if not isinstance(col_data, dict): continue
        for row_idx_str, cell_value in col_data.items():
            if not isinstance(cell_value, str) or not row_idx_str in TX_HEADER_TYPICAL_ROW_INDICES:
                continue
            
            cell_value_stripped = cell_value.strip()
            
            if cell_value_stripped in TRANSACTION_HEADERS_MAP and cell_value_stripped not in located_tx_cols:
                located_tx_cols[TRANSACTION_HEADERS_MAP[cell_value_stripped]] = {"key": col_key, "header_row": row_idx_str, "offset": 1}

            elif cell_value_stripped in MULTI_PART_TX_HEADERS:
                second_part_options = MULTI_PART_TX_HEADERS[cell_value_stripped]
                next_row_idx_str = str(int(row_idx_str) + 1)
                cell_value_next_row = raw_data.get(col_key, {}).get(next_row_idx_str, "").strip()
                for second_part, final_header_key in second_part_options.items():
                    if cell_value_next_row == second_part and final_header_key not in located_tx_cols:
                        located_tx_cols[final_header_key] = {"key": col_key, "header_row": next_row_idx_str, "offset": 1}



    baki_semua_info = located_tx_cols.get("BAKI SEMUA")
    if baki_semua_info:
        baki_semua_col_key = baki_semua_info["key"]
        if baki_semua_col_key.startswith("Unnamed: "):
            try:
                baki_semua_col_num = int(baki_semua_col_key.split(":")[1].strip())
                located_tx_cols["ADDITIONAL_DATA_1"] = {"key": f"Unnamed: {baki_semua_col_num + 1}", "header_row": baki_semua_info["header_row"], "offset": baki_semua_info["offset"]}
                located_tx_cols["ADDITIONAL_DATA_2"] = {"key": f"Unnamed: {baki_semua_col_num + 2}", "header_row": baki_semua_info["header_row"], "offset": baki_semua_info["offset"]}
                located_tx_cols["ADDITIONAL_DATA_3"] = {"key": f"Unnamed: {baki_semua_col_num + 3}", "header_row": baki_semua_info["header_row"], "offset": baki_semua_info["offset"]}
            except (IndexError, ValueError):
                print(f"[WARNING TX] Could not determine additional data columns relative to BAKI SEMUA column '{baki_semua_col_key}'")


    transactions = []
    if not located_tx_cols.get("TARIKH"):
        print("[WARNING TX] TARIKH column for transactions not located. Cannot process transactions.")
        return []

    tarikh_col_info = located_tx_cols["TARIKH"]
    tarikh_data_col_key = tarikh_col_info["key"]
    
    all_header_row_indices = [int(info["header_row"]) for info in located_tx_cols.values() if info.get("header_row")]
    if not all_header_row_indices: 
        print("[WARNING TX] No transaction header row indices found.")
        return []
        
    max_header_row = max(all_header_row_indices)
    data_start_row_num = max_header_row + 1
    
    
    tarikh_data_col_content = raw_data.get(tarikh_data_col_key,{})
    possible_data_row_indices = sorted(
        [k for k in tarikh_data_col_content.keys() if k.isdigit() and int(k) >= data_start_row_num],
        key=int
    )


    for r_idx_str in possible_data_row_indices:
        transaction_item = {}
        has_essential_data = False
        for output_key, header_text in TRANSACTION_HEADERS_MAP.items():
            col_info = located_tx_cols.get(output_key)
            if col_info:
                value = raw_data.get(col_info["key"], {}).get(r_idx_str)
                if output_key == "TARIKH":
                     transaction_item[output_key] = convert_excel_timestamp(value, '%d-%m-%y')
                     if transaction_item[output_key]: has_essential_data = True
                else:
                    transaction_item[output_key] = value


        for add_key in ["ADDITIONAL_DATA_1", "ADDITIONAL_DATA_2", "ADDITIONAL_DATA_3"]:
            col_info = located_tx_cols.get(add_key)
            if col_info:
                value = raw_data.get(col_info["key"], {}).get(r_idx_str)
                if value is not None:
                    transaction_item[add_key] = value
        
        cleaned_transaction_item = {k: v for k, v in transaction_item.items() if v is not None}
        
        if has_essential_data or cleaned_transaction_item.get("PERKARA"):
            transactions.append(cleaned_transaction_item)



    return transactions


def format_dynamically(raw_data):
    statement = {}
    koperasi_col_obj = None
    # Strip KOPERASI_KEY for comparison
    stripped_koperasi_key = KOPERASI_KEY.strip()

    for key, value in raw_data.items():
        if isinstance(key, str) and key.strip() == stripped_koperasi_key:
            koperasi_col_obj = value
            break

    if not koperasi_col_obj:
        print(f"[ERROR] Main Koperasi header column ('{stripped_koperasi_key}') not found in input JSON. Cannot proceed.")
        return {}

    member_data, raw_nominee_data = find_member_data(raw_data, koperasi_col_obj)
    statement.update(member_data)

    nominee_parsed = parse_nominee_string(raw_nominee_data.get("PENAMA_KP_RAW"))
    nominee_output = {
        "NAME": nominee_parsed.get("NAME"),
        "RELATIONSHIP": nominee_parsed.get("RELATIONSHIP"),
        "IC": nominee_parsed.get("IC"),
        "PHONE": raw_nominee_data.get("NO_TELEFON_PENAMA_RAW")
    }
    if any(val for val in nominee_output.values() if val is not None):
        statement["NOMINEE"] = nominee_output
    else:
        statement["NOMINEE"] = None
        
    transactions = find_transaction_columns_and_parse(raw_data, koperasi_col_obj)
    statement["TRANSACTIONS"] = transactions
    
    return statement


def clean_json_data(data):
    if isinstance(data, dict):
        cleaned_dict = {}
        for key, value in data.items():
            if value is None:
                continue
            if (type(value) is int or type(value) is float) and value == 0:
                continue
            cleaned_value = clean_json_data(value)
            if isinstance(cleaned_value, list) and all((type(element) is int or type(element) is float) and element == 0 for element in cleaned_value):
                continue
            if cleaned_value is None and value is not None:
                continue
            cleaned_dict[key] = cleaned_value
        return cleaned_dict
    elif isinstance(data, list):
        cleaned_list = []
        for item in data:
            if item is None:
                continue
            cleaned_item = clean_json_data(item)
            if isinstance(cleaned_item, list) and all((type(element) is int or type(element) is float) and element == 0 for element in cleaned_item):
                continue
            if cleaned_item is None and item is not None:
                continue
            cleaned_list.append(cleaned_item)
        return cleaned_list
    else:
        return data


def excel_to_json(excel_file, sheet_name='Sheet'):
    excel_data_fragment = pd.read_excel(excel_file, sheet_name=sheet_name)
    return excel_data_fragment.to_json()


def main_processing_logic(excel_file, sheet_identifier, output_json):
    try:
        json_str = excel_to_json(excel_file, sheet_identifier)
        json_str = excel_to_json(excel_file, sheet_identifier)
        if not json_str or json_str == "null":
            print(f"Skipping sheet {sheet_identifier} because it's empty or could not be read.")
            return 1
        
        json_data = json.loads(json_str)

        if not json_data:
            print(f"Skipping sheet {sheet_identifier} because it resulted in empty JSON data.")
            return 1

        cleaned_data = clean_json_data(json_data)
        if not cleaned_data:
            print(f"Skipping sheet {sheet_identifier} because cleaned data is empty.")
            return 1

        formatted_data = format_dynamically(cleaned_data)
        
        if not formatted_data:
            print(f"Skipping sheet {sheet_identifier} because formatted data is empty (possibly main Koperasi header was not found or data insufficient).")
            return 1
        
        # Further check if essential member data is present after formatting
        # This helps confirm it's a valid member statement sheet
        if not (formatted_data.get("NO. ANGGOTA") or formatted_data.get("NAMA")):
            print(f"Skipping sheet {sheet_identifier} because essential Koperasi member data (No. Anggota or Nama) is missing after formatting.")
            return 1

        with open(output_json, 'w', encoding='utf-8') as f:
            json.dump(formatted_data, f, indent=4, ensure_ascii=False)
        print(f"Successfully processed sheet {sheet_identifier} to {output_json}.")
        return 0
    except FileNotFoundError:
        print(f"Error: Excel file '{excel_file}' not found for sheet {sheet_identifier}.")
        return 2
    except ValueError as ve:
        print(f"ValueError processing sheet {sheet_identifier}: {ve}")
        return 2
    except KeyError as ke:
        print(f"Skipping sheet {sheet_identifier} as it was not found in the Excel file: {ke}")
        return 1
    except Exception as e:
        print(f"An unexpected error occurred processing sheet {sheet_identifier}: {e}")
        return 2

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Process Excel file and output formatted JSON.")
    parser.add_argument("excel_file", help="Path to the Excel file.")
    parser.add_argument("sheet_name", help="Name of the sheet to convert.")
    parser.add_argument("output_file", help="Path to save the final JSON output.")
    args = parser.parse_args()

    result_code = main_processing_logic(args.excel_file, args.sheet_name, args.output_file)
    sys.exit(result_code)