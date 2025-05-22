import json
import argparse

def is_list_of_only_zeros(value):
    """
    Checks if the given value is a non-empty list containing only the number 0.
    """
    if not isinstance(value, list) or not value:  # Not a list or an empty list
        return False
    # All elements must be the integer 0 or float 0.0
    return all((type(element) is int or type(element) is float) and element == 0 for element in value)

def clean_json_data(data):
    """
    Recursively:
    1. Removes keys from dictionaries where the value is None.
    2. Removes keys from dictionaries where the value is the number 0 (integer or float).
    3. Removes keys from dictionaries where the value, after cleaning, is a list of only zeros.
    4. Removes None items from lists.
    5. Removes items from lists where the item, after cleaning, is a list of only zeros.
    """
    if isinstance(data, dict):
        cleaned_dict = {}
        for key, value in data.items():
            # Rule 1: Skip if original value is None
            if value is None:
                continue
            
            # Rule 2 (New): Skip if original value is the number 0 (int or float)
            # This check ensures we don't remove boolean False.
            if (type(value) is int or type(value) is float) and value == 0:
                continue
            
            cleaned_value = clean_json_data(value)  # Recursively clean the value
            
            # Rule 3: After cleaning, if the value became a list of only zeros, skip this key
            if is_list_of_only_zeros(cleaned_value):
                continue
            
            # Safeguard: if cleaned_value became None after deeper cleaning
            # (e.g., an object that became empty and was set to None - though current logic doesn't make non-None things None)
            if cleaned_value is None and value is not None: 
                 continue

            cleaned_dict[key] = cleaned_value
        return cleaned_dict
    elif isinstance(data, list):
        cleaned_list = []
        for item in data:
            # Rule 4: Skip if original item is None
            if item is None:
                continue
            
            # Rule (Implicit from above): if item is the number 0, it's kept in the list
            # unless the overall list itself becomes a list_of_only_zeros at a higher level.
            # The request "remove these: '216':0" was specific to key-value pairs.
            
            cleaned_item = clean_json_data(item)  # Recursively clean the item
            
            # Rule 5: After cleaning, if the item became a list of only zeros, skip it
            if is_list_of_only_zeros(cleaned_item):
                continue

            # Safeguard if item became None after cleaning deeper structures
            if cleaned_item is None and item is not None:
                 continue
                
            cleaned_list.append(cleaned_item)
        return cleaned_list
    else:
        # Base case: For numbers (non-zero unless in a list), strings, booleans, etc.
        return data

def main():
    parser = argparse.ArgumentParser(
        description="Clean a JSON file by recursively removing: "
                    "1. Keys with 'null' values. "
                    "2. Keys with numerical 0 values (integer or float). "
                    "3. Lists that solely contain numerical 0s (e.g., [0, 0.0]). "
                    "4. 'null' items from lists."
    )
    parser.add_argument("input_file", help="Path to the input JSON file.")
    parser.add_argument(
        "-o", "--output_file", 
        help="Path to the output JSON file. If not provided, prints the cleaned JSON to stdout."
    )
    parser.add_argument(
        "-i", "--indent", 
        type=int, 
        help="Indentation level for the output JSON (e.g., 2 or 4). Omit for compact output."
    )

    args = parser.parse_args()

    try:
        with open(args.input_file, 'r', encoding='utf-8') as f:
            json_data = json.load(f)
    except FileNotFoundError:
        print(f"Error: Input file '{args.input_file}' not found.")
        return
    except json.JSONDecodeError as e:
        print(f"Error: Could not decode JSON from '{args.input_file}'. Invalid JSON: {e}")
        return
    except Exception as e:
        print(f"An error occurred while reading the input file: {e}")
        return

    processed_data = clean_json_data(json_data)

    try:
        if args.output_file:
            with open(args.output_file, 'w', encoding='utf-8') as f:
                json.dump(processed_data, f, indent=args.indent, ensure_ascii=False)
            print(f"Cleaned JSON successfully saved to '{args.output_file}'")
        else:
            # Print to stdout
            print(json.dumps(processed_data, indent=args.indent, ensure_ascii=False))
    except Exception as e:
        print(f"An error occurred while writing the output: {e}")

if __name__ == "__main__":
    main()