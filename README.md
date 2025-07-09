# Excel Data Filtering and Consolidation Script
This Excel filtering code automates user data processing, matching submissions to master records for correction and business unit assignment, and flagging duplicates. It outputs to a multi-sheet Excel, deduplicating valid entries for clean reports while preserving all original invalid records for review.

## What It Does

This tool takes raw user data (e.g., from phishing simulation forms) and processes it to ensure accuracy and consistency. Specifically, it performs the following key functions:

* **Automated Data Matching & Correction:** It intelligently reads user entries, prioritizing email for accurate matching against your master data, then falling back to name. Any discrepancies in user-submitted names, emails, or positions are automatically corrected to align with your master records.
* **Dynamic Business Unit Assignment:** Users are automatically assigned to their respective business units based on their matched master data.
* **Intelligent Duplicate Flagging:** It identifies and flags records that represent consolidated duplicates based on their *corrected* information, helping you understand data quality and potential redundancies.
* **Smart, Conditional Output Organization:** The processed data is exported to a single Excel file containing multiple sheets. Crucially, **valid user data** (matched to a specific Business Unit) is automatically deduplicated for clean reporting, while **all original records** for 'Invalid/Uncategorized' users are preserved on their dedicated sheet, ensuring no data is lost for manual review.

## How It Does It (Technical Logic)

The script operates through several systematic steps:

### 1. Configuration

The script begins by reading configuration settings at the top, including paths to the input Excel file and the directory containing master CSV files, as well as the exact column names expected in both the input and master data. It also allows for mapping internal column names to desired output headers.

### 2. Master Data Loading & Preparation

All `.csv` files within the specified `MASTER_DATA_DIR` are loaded. Each CSV's filename is used to derive and assign a "Business Unit" to the records within that file. The master data (names, emails, positions, companies) is then standardized (e.g., converted to lowercase, extra spaces removed) and indexed into quick lookup maps (by email and name) for efficient matching.

### 3. User Input Processing

The user input Excel file is read, and relevant columns are standardized to prepare for matching. Missing configured input columns are gracefully handled by adding them with `pd.NA`.

### 4. Intelligent Matching & Validation

Each user record is processed through a robust matching logic:

* **Primary Match (Email):** It first attempts to match the user's standardized email against the master data. If a match is found, the user's corrected details (name, email, position, assigned company, and business unit) are populated from the master record. A `Validation_Status` of 'Matched by Email' is assigned.
* **Secondary Match (Name):** If no email match is found, it then attempts to match the user's standardized name. If a name match is successful, but the user's original email differed from the master email found via the name match, the status becomes 'Matched by Name (Email Corrected)'.
* **Invalid/Unmatched:** If neither email nor name matches, the record is flagged as 'Invalid/Unmatched', and corrected fields remain `pd.NA`.

### 5. Duplicate Flagging

After all matching and correction, a `Duplicate_Flag` is assigned to every record. This flag (`'Consolidated Duplicate'` or `'Unique'`) indicates whether the record's combination of corrected business unit, email, name, and position appears more than once *anywhere* in the dataset. This flagging happens *before* any physical removal of duplicates.

### 6. Conditional Output & Deduplication

The final processed data is written to a new Excel file containing multiple sheets:

* **Business Unit Sheets:** For each successfully assigned business unit, a dedicated sheet is created. On these sheets, duplicate records (based on corrected data) are physically removed, keeping only the first occurrence, for a clean, consolidated view.
* **`Invalid_Uncategorized` Sheet:** This sheet contains all records that could not be matched to a master entry. **Crucially, no deduplication is applied here.** All original invalid rows are preserved, even if their corrected (often `pd.NA`) fields are identical, allowing for comprehensive manual review of every unmatched submission.

## User Guide

Follow these steps to set up and run the Excel Data Filtering and Consolidation Script:

### Prerequisites

* **Python 3:** Ensure you have Python 3 installed on your system.
* **Required Libraries:** You will need `pandas` (for data manipulation) and `openpyxl` (for reading and writing Excel files).
    You can install them using pip from your terminal or command prompt:
    ```bash
    pip install pandas openpyxl numpy
    ```

### File Structure

Organize your input Excel file and master data CSVs in a structure similar to this. The script uses relative paths, so ensure your script is placed appropriately.
```bash
├── your_project_folder/
    └── your_script_name.py
    └── excel_file_to_be_filtered.xslx        #data to be cleaned
    └── master_data.cvs                       #all the verified data
```

**Important:** If your file structure differs significantly, you may need to use **absolute paths** (e.g., `C:\Users\YourName\Documents\...`) for the file configuration variables within the script.


### Running the Script

1.  **Open Terminal/Command Prompt:** Navigate to the directory where you saved your `data_filter.py` file.
2.  **Execute Command:** Run the script using the Python interpreter:
    ```bash
    python data_filter.py
    ```
    The script will print its progress and any warnings or errors directly to your console.
