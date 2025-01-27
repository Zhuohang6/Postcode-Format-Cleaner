import pandas as pd
import re
import os
 
# Postcode patterns for all EU countries
EU_POSTCODE_PATTERNS = {
    "AT": r"^\d{4}$",           # Austria
    "BE": r"^\d{4}$",           # Belgium
    "BG": r"^\d{4}$",           # Bulgaria
    "HR": r"^\d{5}$",           # Croatia
    "CY": r"^\d{4}$",           # Cyprus
    "CZ": r"^\d{3} ?\d{2}$",    # Czech Republic
    "DK": r"^\d{4}$",           # Denmark
    "EE": r"^\d{5}$",           # Estonia
    "FI": r"^\d{5}$",           # Finland
    "FR": r"^\d{5}$",           # France
    "DE": r"^\d{5}$",           # Germany
    "GR": r"^\d{3} ?\d{2}$",    # Greece
    "HU": r"^\d{4}$",           # Hungary
    "IE": r"^[A-Z0-9]{3,4} ?[A-Z0-9]{4}$",  # Ireland
    "IT": r"^\d{5}$",           # Italy
    "LV": r"^LV-\d{4}$",        # Latvia
    "LT": r"^LT-\d{5}$",        # Lithuania
    "LU": r"^\d{4}$",           # Luxembourg
    "MT": r"^[A-Z]{3} ?\d{4}$", # Malta
    "NL": r"^\d{4} ?[A-Z]{2}$", # Netherlands
    "PL": r"^\d{2}-\d{3}$",     # Poland
    "PT": r"^\d{4}-\d{3}$",     # Portugal
    "RO": r"^\d{6}$",           # Romania
    "SK": r"^\d{3} ?\d{2}$",    # Slovakia
    "SI": r"^\d{4}$",           # Slovenia
    "ES": r"^\d{5}$",           # Spain
    "SE": r"^\d{3} ?\d{2}$",    # Sweden
    "UK": r"^[A-Z]{1,2}[0-9][0-9A-Z]? ?[0-9][A-Z]{2}$"  # United Kingdom
}
 
def clean_postcodes(postcodes):
    cleaned_postcodes = []
 
    for postcode in postcodes:
        postcode = str(postcode).strip()  # Remove leading/trailing spaces
        postcode = re.sub(r"[^A-Za-z0-9 -]", "", postcode)  # Remove special characters
        postcode = postcode.upper()  # Convert to uppercase
        postcode = re.sub(r"\s+", " ", postcode)  # Normalize spaces
 
        # Pre-validation: Ignore invalid patterns like country names or too-short entries
        if len(postcode) < 5 or not any(char.isdigit() for char in postcode):
            cleaned_postcodes.append({
                'OriginalPostcode': postcode,
                'Country': None,
                'IsValid': False
            })
            continue
 
        # Extract the valid numeric part for validation
        valid_postcode_part = re.match(r"^\d{3} ?\d{2}", postcode)
        if valid_postcode_part:
            postcode = valid_postcode_part.group(0)
 
        # Normalize spaces for UK-style postcodes
        postcode = re.sub(r"\s+", "", postcode)
 
        # Determine country by matching against all patterns
        matched_country = None
        for country, pattern in EU_POSTCODE_PATTERNS.items():
            if re.match(pattern, postcode):
                matched_country = country
                break
 
        # Append result
        cleaned_postcodes.append({
            'OriginalPostcode': postcode,
            'Country': matched_country,
            'IsValid': matched_country is not None
        })
 
    return cleaned_postcodes



if __name__ == "__main__":
    # Specify the input file directly
    input_file = 'your path to addresses.xlsx'
 
    # Validate file existence
    if not os.path.isfile(input_file):
        print(f"Error: File '{input_file}' does not exist.")
        exit(1)
 
    # Determine output file path
    output_file = os.path.splitext(input_file)[0] + "_cleaned.xlsx"
 
    try:
        # Load the Excel file
        df = pd.read_excel(input_file)
 
        if 'PostalCode' not in df.columns:
            print("Error: 'PostalCode' column is missing in the input file.")
            exit(1)
 
        # Clean postcodes
        cleaned_results = clean_postcodes(df['PostalCode'])
 
        # Create a DataFrame from the cleaned results
        cleaned_df = pd.DataFrame(cleaned_results)
 
        # Save to a new Excel file in the same directory as input
        cleaned_df.to_excel(output_file, index=False)
        print(f"Cleaned postcodes saved to {output_file}")
    except Exception as e:
        print(f"An error occurred: {e}")
