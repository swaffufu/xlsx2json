import json
import argparse
from pathlib import Path
from datetime import datetime, timedelta # Added timedelta for Excel serial dates
import re

# --- Constants for known labels and headers ---
# Key for the column often containing member labels and transaction dates
KOPERASI_KEY = "              KOPERASI PERMODALAN DAN PERUSAHAAN MELAYU NEGERI SEMBILAN BERHAD" # Note leading space

# Member labels we expect to find and their desired output keys
MEMBER_LABELS_MAP = {
    "NO. ANGGOTA": "NO. ANGGOTA", "GELARAN": "GELARAN", "NAMA": "NAMA",
    "NO. K/P": "NO. K/P", "TARIKH LAHIR": "TARIKH LAHIR", "ALAMAT TETAP": "ALAMAT TETAP",
    "ALAMAT SURAT MENYURAT": "ALAMAT SURAT MENYURAT", 
    "NO. TELEFON ANGGOTA": "NO. TELEFON ANGGOTA", "PERKERJAAN": "PERKERJAAN",
    "PENAMA / K.P": "PENAMA_KP_RAW", # Will be parsed into a NOMINEE object
    "NO. TELEFON PENAMA ": "NO_TELEFON_PENAMA_RAW", # Note the trailing space, as seen in JSON sample
    "TARIKH MASUK": "TARIKH MASUK", "TARIKH LULUS ALK": "TARIKH LULUS ALK"
}
# Expected row indices (as strings) for these labels within the KOPERASI_KEY object
MEMBER_LABELS_EXPECTED_ROW_INDICES = {
    "NO. ANGGOTA": "5", "GELARAN": "6", "NAMA": "7", "NO. K/P": "8",
    "TARIKH LAHIR": "9", "ALAMAT TETAP": "10", "ALAMAT SURAT MENYURAT": "11",
    "NO. TELEFON ANGGOTA": "12", "PERKERJAAN": "13", "PENAMA / K.P": "14",
    "NO. TELEFON PENAMA ": "15", "TARIKH MASUK": "16", "TARIKH LULUS ALK": "17"
}

# Transaction headers to search for and their desired output keys
TRANSACTION_HEADERS_MAP = {
    "TARIKH": "TARIKH", "PERKARA": "PERKARA", "NO.RESIT": "NO.RESIT", "TAHUN": "TAHUN",
    "WANG MASUK": "WANG MASUK", "WANG KELUAR": "WANG KELUAR",
    "BAKI SYER": "BAKI SYER", "BAKI BONUS": "BAKI BONUS", "BAKI SEMUA": "BAKI SEMUA",
    "CATATAN": "CATATAN"
}
# For multi-part headers, map the first part to the second part and then to the final key
MULTI_PART_TX_HEADERS = {
    "WANG": {"MASUK": "WANG MASUK"},
    "BAKI": {"SYER": "BAKI SYER", "BONUS": "BAKI BONUS", "SEMUA": "BAKI SEMUA"}
}
TX_HEADER_TYPICAL_ROW_INDICES = ("19", "20") # Headers often on these row indices

# --- Utility Functions ---
def convert_excel_timestamp(value, date_format='%d-%b-%y'):
    if value is None: return None
    if isinstance(value, str):
        try:
            common_formats = ['%d-%m-%y', '%d/%m/%y', '%d-%b-%y', '%Y-%m-%d', '%m/%d/%Y', '%d-%m-%Y', '%d/%m/%Y']
            dt_object = None
            for fmt in common_formats:
                try:
                    dt_object = datetime.strptime(value, fmt)
                    break
                except ValueError:
                    continue
            if dt_object:
                return dt_object.strftime(date_format).upper()
            return value.upper() 
        except ValueError:
            return value.upper()
            
    if isinstance(value, (int, float)):
        try:
            # Heuristic for Unix millisecond timestamps
            if value > 10**11 and value < 4 * 10**12 : # Plausible range for ms timestamps around current era
                dt_object = datetime.fromtimestamp(value / 1000)
            # Heuristic for Excel serial date numbers (days since 1899-12-30 for Windows default)
            elif value > 10000 and value < 70000: # Approx range for 20th/21st century dates
                dt_object = datetime(1899, 12, 30) + timedelta(days=value)
            else: # If not clearly identifiable as a common timestamp type
                return str(int(value)) if isinstance(value, float) and value == int(value) else str(value)
            return dt_object.strftime(date_format).upper()
        except (ValueError, TypeError, OSError): # OSError for out-of-range timestamps
            return str(value)
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
    
    # 1. Attempt to identify the primary Member Value Column
    member_value_col_key_found = None
    no_anggota_label_row = MEMBER_LABELS_EXPECTED_ROW_INDICES.get("NO. ANGGOTA")
    
    if no_anggota_label_row and koperasi_col_obj.get(no_anggota_label_row) == "NO. ANGGOTA":
        # Search horizontally for an integer value that could be the member number
        # Check same row, row below (for json_5000 anomaly where value was at label_row+1)
        # and row above (for json_5000 anomaly where GELARAN value was at label_row-1)
        row_indices_to_check_for_value = [
            no_anggota_label_row,
            str(int(no_anggota_label_row) + 1),
            str(int(no_anggota_label_row) - 1) 
        ]
        # Prefer "Unnamed: 3" or "Unnamed: 4" if member number is found there
        preferred_value_cols = ["Unnamed: 3", "Unnamed: 4"] 
        other_potential_value_cols = [k for k in raw_data.keys() if k.startswith("Unnamed:") and k not in preferred_value_cols]
        
        for r_idx_val_check in row_indices_to_check_for_value:
            for col_key_cand in preferred_value_cols + other_potential_value_cols:
                col_obj_cand = raw_data.get(col_key_cand, {})
                val = col_obj_cand.get(r_idx_val_check)
                if isinstance(val, int) and 1000 <= val <= 99999: # Plausible member no.
                    member_value_col_key_found = col_key_cand
                    # Store actual row index where NO. ANGGOTA value was found
                    member_details[MEMBER_LABELS_MAP["NO. ANGGOTA"]] = val
                    # This is a bit heuristic, assumes NO. ANGGOTA is the first reliable find
                    # We'll store this found value's row index for potential relative lookups
                    no_anggota_value_row_idx = r_idx_val_check 
                    break
            if member_value_col_key_found:
                break
    
    if not member_value_col_key_found:
        print("[WARNING] Could not dynamically identify the primary Member Value Column. Member details might be incomplete.")
        # As a last resort, try the columns used in previous examples if all else fails
        if raw_data.get("Unnamed: 4", {}).get(MEMBER_LABELS_EXPECTED_ROW_INDICES.get("NAMA")): # Check if "Unnamed: 4" looks like a value col
            member_value_col_key_found = "Unnamed: 4"
        elif raw_data.get("Unnamed: 3", {}).get(MEMBER_LABELS_EXPECTED_ROW_INDICES.get("NAMA")):
             member_value_col_key_found = "Unnamed: 3"


    # 2. Extract other member details using the identified (or fallback) value column
    member_value_col_data = raw_data.get(member_value_col_key_found, {}) if member_value_col_key_found else {}
    print(f"[DEBUG MEMBER] Using member value column: '{member_value_col_key_found}' with data: {str(list(member_value_col_data.items())[:5])[:100]}...")


    for label_text, output_key in MEMBER_LABELS_MAP.items():
        if output_key in member_details: # Already found (like NO. ANGGOTA)
            continue

        expected_label_row_idx = MEMBER_LABELS_EXPECTED_ROW_INDICES.get(label_text)
        if not expected_label_row_idx or koperasi_col_obj.get(expected_label_row_idx) != label_text:
            # print(f"[DEBUG MEMBER] Label '{label_text}' not found at expected row '{expected_label_row_idx}' in KOPERASI_KEY column.")
            continue

        # Default: try to get value from the same row index as the label
        value = member_value_col_data.get(expected_label_row_idx)

        # Specific heuristic for GELARAN based on json_5000 structure if direct match fails
        # (where NO.ANGGOTA label was "5", value in "Unnamed:4" was at "6";
        #  GELARAN label was "6", value in "Unnamed:4" was at "5")
        if label_text == "GELARAN" and value is None and member_value_col_key_found == "Unnamed: 4": # Typical for json_5000
             # If NO.ANGGOTA value was found at its label_row+1, GELARAN value might be at its label_row-1
            if no_anggota_value_row_idx == str(int(MEMBER_LABELS_EXPECTED_ROW_INDICES["NO. ANGGOTA"]) + 1) :
                value = member_value_col_data.get(str(int(expected_label_row_idx) - 1))


        if output_key == "PENAMA_KP_RAW":
            raw_nominee_details["PENAMA_KP_RAW"] = value
        elif output_key == "NO_TELEFON_PENAMA_RAW":
            raw_nominee_details["NO_TELEFON_PENAMA_RAW"] = value
        elif "TARIKH" in output_key: # TARIKH LAHIR, MASUK, LULUS ALK
            member_details[output_key] = convert_excel_timestamp(value, '%d-%b-%y')
        else:
            member_details[output_key] = value
            
    return member_details, raw_nominee_details


def find_transaction_columns_and_parse(raw_data, koperasi_col_obj):
    located_tx_cols = {} # Map: "TARIKH" -> {"col_key": "...", "header_row": "...", "data_start_row_offset": 1 or 2}
    
    print("[DEBUG TX] Locating transaction headers...")
    for col_key, col_data in raw_data.items():
        if not isinstance(col_data, dict): continue
        for row_idx_str, cell_value in col_data.items():
            if not isinstance(cell_value, str) or not row_idx_str in TX_HEADER_TYPICAL_ROW_INDICES:
                continue
            
            cell_value_stripped = cell_value.strip()
            
            # Check single line headers
            if cell_value_stripped in TRANSACTION_HEADERS_MAP and cell_value_stripped not in located_tx_cols:
                located_tx_cols[TRANSACTION_HEADERS_MAP[cell_value_stripped]] = {"key": col_key, "header_row": row_idx_str, "offset": 1}
                # print(f"[DEBUG TX] Found header '{cell_value_stripped}' in col '{col_key}' at row '{row_idx_str}'")

            # Check multi-part headers (e.g., "WANG" then "MASUK")
            elif cell_value_stripped in MULTI_PART_TX_HEADERS:
                second_part_options = MULTI_PART_TX_HEADERS[cell_value_stripped]
                next_row_idx_str = str(int(row_idx_str) + 1)
                cell_value_next_row = raw_data.get(col_key, {}).get(next_row_idx_str, "").strip()
                for second_part, final_header_key in second_part_options.items():
                    if cell_value_next_row == second_part and final_header_key not in located_tx_cols:
                        located_tx_cols[final_header_key] = {"key": col_key, "header_row": next_row_idx_str, "offset": 1} # Data starts 1 row after the second part
                        # print(f"[DEBUG TX] Found multi-part header '{final_header_key}' in col '{col_key}', part2 at row '{next_row_idx_str}'")


    # Identify columns for "ADDITIONAL_DATA_1, 2, 3" relative to "BAKI SEMUA"
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
    print(f"[DEBUG TX] Located transaction columns map: {json.dumps(located_tx_cols, indent=2)}")

    # --- Extract Transaction Rows ---
    transactions = []
    if not located_tx_cols.get("TARIKH"):
        print("[WARNING TX] TARIKH column for transactions not located. Cannot process transactions.")
        return []

    tarikh_col_info = located_tx_cols["TARIKH"]
    tarikh_data_col_key = tarikh_col_info["key"]
    # Determine data start row: typically the row after all headers are defined.
    # For simplicity, let's assume data starts 1 or 2 rows after the "TARIKH" header.
    # A more robust way is to find the max header row index across all located headers.
    
    all_header_row_indices = [int(info["header_row"]) for info in located_tx_cols.values() if info.get("header_row")]
    if not all_header_row_indices: 
        print("[WARNING TX] No transaction header row indices found.")
        return []
        
    max_header_row = max(all_header_row_indices)
    data_start_row_num = max_header_row + 1 # Start looking for data from here
    
    print(f"[DEBUG TX] Max header row index found: {max_header_row}. Data rows start from index {data_start_row_num}.")

    # Iterate through potential data row indices
    # Get all numeric row keys from the TARIKH data column that are >= data_start_row_num
    
    tarikh_data_col_content = raw_data.get(tarikh_data_col_key,{})
    possible_data_row_indices = sorted(
        [k for k in tarikh_data_col_content.keys() if k.isdigit() and int(k) >= data_start_row_num],
        key=int
    )
    print(f"[DEBUG TX] Potential data row indices for transactions: {possible_data_row_indices}")


    for r_idx_str in possible_data_row_indices:
        transaction_item = {}
        has_essential_data = False
        for output_key, header_text in TRANSACTION_HEADERS_MAP.items(): # Use original map to iterate desired fields
            col_info = located_tx_cols.get(output_key)
            if col_info:
                value = raw_data.get(col_info["key"], {}).get(r_idx_str)
                if output_key == "TARIKH": # Transaction TARIKH format DD-MM-YY
                     transaction_item[output_key] = convert_excel_timestamp(value, '%d-%m-%y')
                     if transaction_item[output_key]: has_essential_data = True
                else:
                    transaction_item[output_key] = value
            # else:
                # print(f"[DEBUG TX] Column info for header '{output_key}' not found for row '{r_idx_str}'.")

        # Add ADDITIONAL_DATA if their columns were mapped
        for add_key in ["ADDITIONAL_DATA_1", "ADDITIONAL_DATA_2", "ADDITIONAL_DATA_3"]:
            col_info = located_tx_cols.get(add_key)
            if col_info:
                value = raw_data.get(col_info["key"], {}).get(r_idx_str)
                if value is not None: # Only add if value exists
                    transaction_item[add_key] = value
        
        cleaned_transaction_item = {k: v for k, v in transaction_item.items() if v is not None}
        
        if has_essential_data or cleaned_transaction_item.get("PERKARA"): # Add if TARIKH or PERKARA is present
            transactions.append(cleaned_transaction_item)
            # print(f"[DEBUG TX] Added transaction for row index '{r_idx_str}': {cleaned_transaction_item}")
        # else:
            # print(f"[DEBUG TX] Skipped transaction for row index '{r_idx_str}' due to missing essential data: {cleaned_transaction_item}")


    return transactions


def format_dynamically(raw_data):
    statement = {}
    koperasi_col_obj = raw_data.get(KOPERASI_KEY, {})

    if not koperasi_col_obj:
        print(f"[ERROR] Main Koperasi header column ('{KOPERASI_KEY}') not found in input JSON. Cannot proceed.")
        return {}

    # 1. Member and Raw Nominee Details
    member_data, raw_nominee_data = find_member_data(raw_data, koperasi_col_obj)
    statement.update(member_data)

    # 2. Parse and Add Nominee Object
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
        
    # 3. Transactions
    transactions = find_transaction_columns_and_parse(raw_data, koperasi_col_obj)
    statement["TRANSACTIONS"] = transactions
    
    return statement

def main():
    parser = argparse.ArgumentParser(
        description="Dynamically parses a pandas-generated JSON and formats it into a nested Koperasi statement."
    )
    parser.add_argument("input_file", help="Path to the input JSON file (pandas-like structure).")
    parser.add_argument("-o", "--output_file", required=True, help="Path to save the formatted JSON output.")
    parser.add_argument("-i", "--indent", type=int, default=2, help="Indentation for output JSON (default: 2).")

    args = parser.parse_args()

    try:
        with open(args.input_file, 'r', encoding='utf-8') as f:
            raw_json_data = json.load(f)
    except FileNotFoundError:
        print(f"Error: Input file '{args.input_file}' not found.")
        return
    except json.JSONDecodeError as e:
        print(f"Error: Could not decode JSON from '{args.input_file}'. Invalid JSON: {e}")
        return
    except Exception as e:
        print(f"An unexpected error occurred while reading input: {e}")
        return

    formatted_data = format_dynamically(raw_json_data)

    if not formatted_data or not formatted_data.get("NO. ANGGOTA"):
        print("Error: Failed to extract sufficient data to produce a formatted statement. Output might be empty or incomplete.")
    
    try:
        with open(args.output_file, 'w', encoding='utf-8') as f:
            json.dump(formatted_data, f, indent=args.indent, ensure_ascii=False)
        print(f"Formatted JSON successfully saved to '{args.output_file}'")
    except Exception as e:
        print(f"An error occurred while writing output: {e}")

if __name__ == "__main__":
    main()