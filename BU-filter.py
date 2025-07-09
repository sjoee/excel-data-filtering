import pandas as pd
import numpy as np
import os # For file system operations

# --- Configuration ---
# Use raw string literal (r"...") for Windows paths to avoid issues with backslashes
# **CRITICAL: Ensure these paths are EXACTLY as they appear in your File Explorer**
USER_INPUT_FILE = r'.\excel_file_to_be_filtered.xlsx'
USER_INPUT_SHEET = 'Sheet1'

MASTER_DATA_DIR = r'.\master_data.cvs'

# --- Dynamic Output File Naming ---
input_filename = os.path.basename(USER_INPUT_FILE)
input_file_name_without_ext, _ = os.path.splitext(input_filename) # Use _ to discard extension if not needed later

# Construct the new output filename using the input file's name
OUTPUT_FILTERED_FILE = f'filtered-{input_file_name_without_ext}.xlsx'

# Columns in user input form (ensure these match your Excel sheet headers)
COL_INPUT_NAME = 'Full Name (as per NRIC/Passport)'
COL_INPUT_EMAIL = 'Work Email Address'
COL_INPUT_POSITION = 'Position / Job Title'
COL_INPUT_BU_RAW = 'Department / Business Unit' # User-entered business unit (will likely be ignored if Company_filter_words are used primarily)

# Columns expected in each master CSV file (ensure these match your CSV headers exactly)
COL_MASTER_USER_NAME = 'First Name'
COL_MASTER_USER_EMAIL = 'Email'
COL_MASTER_USER_POSITION = 'Position'
COL_MASTER_COMPANY = 'Company' 

# The name for the Business Unit column we will *add* to the master DataFrame
NEW_COL_MASTER_BU = 'Business Unit' # This will hold the BU name derived from CSV filename

# Define your desired output column headers here
# Map 'Internal_DataFrame_Column_Name': 'Desired Output Header Name'
# This dictionary will be used to rename columns just before writing to Excel.
OUTPUT_COLUMN_HEADERS_MAP = {
    'Corrected_Name': 'Full Name',
    'Corrected_Position': 'Position',
    'Assigned_Company': 'Company',
    'Corrected_Email': 'Email',
    COL_INPUT_NAME: 'Original Input Name',
    COL_INPUT_EMAIL: 'Original Input Email',
    COL_INPUT_POSITION: 'Original Input Position',
    COL_INPUT_BU_RAW: 'Original Department/BU',
    'Assigned_Business_Unit': 'Assigned Business Unit',
    'Validation_Status': 'Validation Status',
    'Duplicate_Flag': 'Duplicate Status' # Changed to 'Duplicate Status' for clarity with new logic
}

# --- Helper Function for Standardization (Crucial for matching) ---
def standardize_text(text):
    if pd.isna(text):
        return np.nan
    s_text = str(text).strip().lower()
    # Add/adjust these replacements based on common variations in your data
    s_text = s_text.replace('.', '').replace(' ', '').replace('grp', 'group').replace('dept', '')
    return s_text

# --- Function to Load Master Data from Multiple CSVs ---
def load_master_data_from_csvs(directory_path, actual_bu_column_name_to_add, expected_cols_in_csv):
    all_master_dfs = []
    print(f"Loading master data from CSVs in directory: '{directory_path}'...")
    if not os.path.exists(directory_path):
        raise FileNotFoundError(f"Master data directory '{directory_path}' not found at: '{os.path.abspath(directory_path)}'")

    for filename in os.listdir(directory_path):
        if filename.endswith('.csv'):
            filepath = os.path.join(directory_path, filename)
            bu_name = os.path.splitext(filename)[0] # Extract BU name from filename (e.g., 'Sales' from 'Sales.csv')
            print(f"  Processing '{filename}' (Derived Business Unit: '{bu_name}')")
            try:
                # Specify usecols to only read necessary columns, improving performance and avoiding warnings
                df = pd.read_csv(filepath, usecols=expected_cols_in_csv)

                # Validate essential columns exist in the loaded CSV
                missing_cols = [col for col in expected_cols_in_csv if col not in df.columns]
                if missing_cols:
                    print(f"  Warning: Skipping '{filename}' due to missing essential columns: {missing_cols}")
                    continue

                # Add the Business Unit column based on the filename
                df[actual_bu_column_name_to_add] = bu_name

                all_master_dfs.append(df)
                print(f"    Loaded {len(df)} rows.")
            except pd.errors.EmptyDataError:
                print(f"  Warning: Skipping empty CSV file '{filename}'.")
            except ValueError as ve: # Catches errors if usecols specifies non-existent column
                print(f"  Error loading '{filename}' with specified columns: {ve}")
                print(f"  Please ensure all columns in `expected_cols_in_csv` ({expected_cols_in_csv}) exist in '{filename}'.")
                continue
            except Exception as e:
                print(f"  Error loading '{filename}': {e}")

    if not all_master_dfs:
        raise ValueError(f"No valid CSV files found or loaded from '{directory_path}'. Please check directory and file contents.")

    # Concatenate all DataFrames into a single master DataFrame
    master_df = pd.concat(all_master_dfs, ignore_index=True)
    return master_df

# --- Main Script ---
try:
    print(f"Attempting to load user input from: '{USER_INPUT_FILE}'...")
    # Add an explicit check and print the absolute path Python is looking for
    if not os.path.exists(USER_INPUT_FILE):
        print(f"ERROR: The user input file was not found at the specified path.")
        print(f"Absolute path attempted: '{os.path.abspath(USER_INPUT_FILE)}'")
        raise FileNotFoundError(f"Input file not found: '{USER_INPUT_FILE}'")

    # List all potential input columns that the script is configured to look for
    all_potential_input_columns = [
        COL_INPUT_NAME,
        COL_INPUT_EMAIL,
        COL_INPUT_POSITION,
        COL_INPUT_BU_RAW
    ]
    
    # Read just the header row to get actual column names in the Excel file
    excel_cols = pd.read_excel(USER_INPUT_FILE, sheet_name=USER_INPUT_SHEET, nrows=0).columns.tolist()
    
    # Determine which configured columns are actually present in the Excel file
    actual_cols_to_read = [col for col in all_potential_input_columns if col in excel_cols]

    if not actual_cols_to_read:
        # If no configured columns are found in the Excel, this indicates a major configuration issue
        raise ValueError(f"No usable columns found in input Excel sheet '{USER_INPUT_SHEET}' that match configured COL_INPUT_... variables. Please check your Excel headers and script configuration.")
    
    # Inform the user about any configured columns that were not found and thus will be skipped
    missing_but_skipped_cols = [col for col in all_potential_input_columns if col not in excel_cols]
    if missing_but_skipped_cols:
        print(f"Warning: The following input columns were configured but not found in '{USER_INPUT_FILE}' and will be treated as empty: {missing_but_skipped_cols}")

    # Read the input Excel file, only loading the columns that actually exist
    user_input_df = pd.read_excel(
        USER_INPUT_FILE,
        sheet_name=USER_INPUT_SHEET,
        usecols=actual_cols_to_read 
    )
    
    # For any configured COL_INPUT_ columns that were *not* present in the Excel,
    # add them to the DataFrame now, filled with NA values. This prevents KeyErrors later.
    for col in all_potential_input_columns:
        if col not in user_input_df.columns:
            user_input_df[col] = pd.NA

    # 1. Load Master Data from CSVs
    expected_cols_in_each_csv = [
        COL_MASTER_USER_NAME,
        COL_MASTER_USER_EMAIL,
        COL_MASTER_USER_POSITION,
        COL_MASTER_COMPANY
    ]
    master_df = load_master_data_from_csvs(MASTER_DATA_DIR, NEW_COL_MASTER_BU, expected_cols_in_each_csv)

    print("\n--- Consolidated Master Companies & Business Units ---")
    print(f"Total master records loaded: {len(master_df)}")

    # --- 2. Prepare Master Data for Lookups ---
    master_df['Std_Master_Company'] = master_df[COL_MASTER_COMPANY].apply(standardize_text)
    master_df['Std_Master_BU'] = master_df[NEW_COL_MASTER_BU].apply(standardize_text)
    master_df['Std_Master_Email'] = master_df[COL_MASTER_USER_EMAIL].apply(standardize_text)
    master_df['Std_Master_User_Name'] = master_df[COL_MASTER_USER_NAME].apply(standardize_text)

    # Create maps for detailed master record lookup
    email_to_master_details_map = {}
    name_to_master_details_map = {} 

    for idx, row in master_df.iterrows():
        master_details = {
            'actual_company': row[COL_MASTER_COMPANY],
            'actual_bu': row[NEW_COL_MASTER_BU],
            'master_name': row[COL_MASTER_USER_NAME],
            'master_email': row[COL_MASTER_USER_EMAIL],
            'master_position': row[COL_MASTER_USER_POSITION]
        }
        if pd.notna(row['Std_Master_Email']):
            email_to_master_details_map[row['Std_Master_Email']] = master_details
        if pd.notna(row['Std_Master_User_Name']):
            # If names are not unique in master data, this will store the last encountered one.
            # You might want to refine this if names are not unique identifiers in your master data.
            name_to_master_details_map[row['Std_Master_User_Name']] = master_details

    print("\nMaster Data Prepared for Lookups.")

    # --- 3. Prepare User Input Data ---
    user_input_df['Std_Input_Email'] = user_input_df[COL_INPUT_EMAIL].apply(standardize_text)
    user_input_df['Std_Input_Name'] = user_input_df[COL_INPUT_NAME].apply(standardize_text)
    user_input_df['Std_Input_BU'] = user_input_df[COL_INPUT_BU_RAW].apply(standardize_text)

    # --- 4. Match and Validate User Input against Master Data with Correction ---

    user_input_df['Assigned_Company'] = pd.NA
    user_input_df['Assigned_Business_Unit'] = pd.NA
    user_input_df['Corrected_Name'] = pd.NA
    user_input_df['Corrected_Email'] = pd.NA
    user_input_df['Corrected_Position'] = pd.NA
    user_input_df['Validation_Status'] = 'Unprocessed'

    for index, row in user_input_df.iterrows():
        std_email = row['Std_Input_Email']
        std_name = row['Std_Input_Name']

        matched_details = None
        status = 'Invalid/Unmatched'

        # Attempt 1: Match by Email (Highest Priority for specific user details)
        if pd.notna(std_email) and std_email in email_to_master_details_map:
            matched_details = email_to_master_details_map[std_email]
            status = 'Matched by Email'
            # Optional: Add this print to see name corrections happening
            if pd.notna(row[COL_INPUT_NAME]) and pd.notna(matched_details['master_name']) and \
               standardize_text(row[COL_INPUT_NAME]) != standardize_text(matched_details['master_name']):
                print(f"  Info: Correcting name for email '{row[COL_INPUT_EMAIL]}': '{row[COL_INPUT_NAME]}' -> '{matched_details['master_name']}'")
        # Attempt 2: If Email Fails, Match by Name (Secondary Priority for specific user details)
        elif pd.notna(std_name) and std_name in name_to_master_details_map:
            matched_details = name_to_master_details_map[std_name]
            status = 'Matched by Name'
            # Flag if email in input differs from master email found by name
            if pd.notna(std_email) and pd.notna(matched_details['master_email']) and std_email != standardize_text(matched_details['master_email']):
                status = 'Matched by Name (Email Corrected)'

        # Apply matched details if found
        if matched_details:
            user_input_df.at[index, 'Assigned_Company'] = matched_details['actual_company']
            user_input_df.at[index, 'Assigned_Business_Unit'] = matched_details['actual_bu']
            user_input_df.at[index, 'Corrected_Name'] = matched_details['master_name']
            user_input_df.at[index, 'Corrected_Email'] = matched_details['master_email']
            user_input_df.at[index, 'Corrected_Position'] = matched_details['master_position'] 
            user_input_df.at[index, 'Validation_Status'] = status
        else:
            # If no specific user match (by email or name),
            # it's considered Invalid/Unmatched by default.
            user_input_df.at[index, 'Validation_Status'] = 'Invalid/Unmatched'
            user_input_df.at[index, 'Corrected_Name'] = pd.NA
            user_input_df.at[index, 'Corrected_Email'] = pd.NA
            user_input_df.at[index, 'Corrected_Position'] = pd.NA

    # Filter out temporary columns before final output
    final_df = user_input_df.drop(columns=[
        'Std_Input_Email', 'Std_Input_Name', 'Std_Input_BU'
    ])

    # --- 5. Flag Duplicate Data (Flagging only, no global deduplication) ---
    # Define columns to check for duplicates among the *corrected* values
    # These are the columns that define a "unique" person for consolidation if it were applied
    duplicate_check_cols_for_consolidation = [
        'Assigned_Business_Unit',
        'Corrected_Email',
        'Corrected_Name',
        'Corrected_Position'
    ]
    
    # Ensure these columns actually exist in final_df before using them
    actual_cols_for_dup_check = [col for col in duplicate_check_cols_for_consolidation if col in final_df.columns]

    if not final_df.empty and actual_cols_for_dup_check:
        # Identify groups that have duplicates (keep=False marks ALL duplicates in a group)
        # This column now indicates if a row *is part of a duplicate group* based on its corrected info
        final_df['_is_part_of_duplicate_group'] = final_df.duplicated(subset=actual_cols_for_dup_check, keep=False)

        # Assign Duplicate_Flag based on whether the record was part of a duplicate group
        final_df['Duplicate_Flag'] = final_df['_is_part_of_duplicate_group'].apply(
            lambda x: 'Consolidated Duplicate' if x else 'Unique'
        )
        
        # Drop the temporary column
        final_df = final_df.drop(columns=['_is_part_of_duplicate_group'])

        print(f"\n--- Duplicate Flagging Complete (Flags assigned, no global deduplication) ---")
    else:
        final_df['Duplicate_Flag'] = 'Unique' # If no data or no columns to check, all are unique

    print("\n--- Processed User Data (with Duplicate Flag) ---")
    print(final_df.head(10)) 
    print("\nValidation Status Summary:")
    print(final_df['Validation_Status'].value_counts())
    print("\nDuplicate Flag Summary:")
    print(final_df['Duplicate_Flag'].value_counts(dropna=False))

    # --- 6. Prepare for Output: Group by Assigned_Business_Unit ---
    # Get unique BU values, including pd.NA for the 'Invalid_Uncategorized' sheet
    assigned_bus = final_df['Assigned_Business_Unit'].astype(str).unique()
    print(f"\nUnique Assigned Business Units found for output: {assigned_bus}")

    output_filepath = OUTPUT_FILTERED_FILE
    with pd.ExcelWriter(output_filepath, engine='openpyxl') as writer:
        for bu_val in assigned_bus:
            # Determine if this is the 'Invalid_Uncategorized' sheet
            is_invalid_sheet = (bu_val == str(pd.NA))

            if is_invalid_sheet:
                sheet_name = 'Invalid_Uncategorized'
                subset_df = final_df[final_df['Assigned_Business_Unit'].isna()].copy()

                # --- Specific Headers for Invalid_Uncategorized Sheet (All original & corrected) ---
                internal_cols_to_output_order = [
                    COL_INPUT_NAME,         # Original Input Name
                    COL_INPUT_EMAIL,        # Original Input Email
                    COL_INPUT_POSITION,     # Original Input Position
                    COL_INPUT_BU_RAW,       # Original Input Business Unit
                    'Corrected_Name',
                    'Corrected_Position',
                    'Assigned_Company',
                    'Corrected_Email',
                    'Assigned_Business_Unit', # Will be NA for this sheet, but still useful to show
                    'Validation_Status'
                ]
                # No deduplication for the invalid sheet: keep all original rows
                # (even if their corrected data is identical like all NAs)

            else: # This is a valid Business Unit sheet
                # Sanitize sheet name for Excel (max 31 chars, no invalid chars)
                sheet_name = str(bu_val).replace('/', '-').replace('\\', '-').replace(':', '').replace('*', '').replace('?', '').replace('[', '').replace(']', '')
                if len(sheet_name) > 31:
                    sheet_name = sheet_name[:31] # Truncate if too long
                subset_df = final_df[final_df['Assigned_Business_Unit'] == bu_val].copy()

                # --- Standard Headers for Valid BU-specific Sheets ---
                internal_cols_to_output_order = [
                    'Corrected_Name',
                    'Corrected_Position',
                    'Assigned_Company',
                    'Corrected_Email',
                    COL_INPUT_BU_RAW,       # Original Input Business Unit (still useful for BU sheets)
                    'Assigned_Business_Unit',
                    'Validation_Status',
                    'Duplicate_Flag'
                ]
                
                # --- APPLY DEDUPLICATION FOR VALID BU SHEETS ONLY ---
                # This ensures only unique corrected records are on the valid BU sheets
                if not subset_df.empty and actual_cols_for_dup_check:
                    original_len = len(subset_df)
                    subset_df.drop_duplicates(subset=actual_cols_for_dup_check, keep='first', inplace=True)
                    if len(subset_df) < original_len:
                        print(f"    Deduplicated {original_len - len(subset_df)} records for sheet: '{sheet_name}'")


            # Filter output columns to only include those actually present in the DataFrame
            existing_cols_for_output = [col for col in internal_cols_to_output_order if col in subset_df.columns]

            if not subset_df.empty:
                print(f"Writing {len(subset_df)} records to sheet: '{sheet_name}'")
                
                # Select, order, and then rename columns for the Excel output
                df_for_excel_output = subset_df[existing_cols_for_output].rename(columns=OUTPUT_COLUMN_HEADERS_MAP)
                
                # Write the prepared DataFrame to the Excel sheet
                df_for_excel_output.to_excel(writer, sheet_name=sheet_name, index=False)
            else:
                print(f"No records for Business Unit: '{bu_val}' - Skipping sheet creation.")

    print(f"\nSuccessfully filtered and saved data to '{output_filepath}' by Business Unit.")

except FileNotFoundError as e:
    print(f"ERROR: A required file or directory was not found. Please check paths. {e}")
    print(f"Ensure that '{USER_INPUT_FILE}' exists at the specified location and that the directory '{MASTER_DATA_DIR}' exists and contains CSVs.")
    print(f"If the file is actually named differently, please update USER_INPUT_FILE and MASTER_DATA_DIR variables with exact paths and names.")
    print(f"For absolute paths (like yours), ensure every segment matches your file system exactly.")
except KeyError as e:
    print(f"ERROR: Missing expected column. This might be a configuration issue. Column: {e}")
    print(f"Please verify that the column names configured in COL_INPUT_ and COL_MASTER_ variables exactly match the headers in your input Excel and master CSV files.")
    print(f"If you see this error for a column that should be in a master CSV, ensure the CSV itself is not malformed or truly missing that column.")
except ValueError as e:
    print(f"ERROR: A data processing issue occurred: {e}")
    print(f"This often indicates a problem with data consistency or file structure. Review the error message for specifics.")
except Exception as e:
    print(f"An unexpected error occurred: {e}")
    import traceback
    traceback.print_exc()
