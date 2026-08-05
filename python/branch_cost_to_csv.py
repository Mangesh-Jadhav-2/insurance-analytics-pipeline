"""
branch_cost_to_csv.py

Converts a branch operating-expense Excel report into a clean CSV ready for
loading into a SQL table: maps human-readable Excel column headers to
snake_case SQL column names, coerces data types (strings, ints, floats,
dates), and fills in any missing expected columns with sane defaults.

Usage:
    python branch_cost_to_csv.py --input path/to/report.xlsx --output path/to/report.csv
"""

import argparse

import pandas as pd

# Excel header -> SQL table column name
COLUMN_MAPPING = {
    'RegionName': 'region_name',
    'Branch Name/Code': 'branch_name',
    'Branch_name': 'actual_branch',
    'Actual Branch Name': 'branch_location',
    'Branch manager': 'branch_manager',
    'RH Name': 'rh_name_name',
    'No. of Employees': 'no_of_emp',
    'Month': 'month_period',
    'Courier Charges': 'courier_charges',
    'Internet Charges': 'internet_charges',
    'Telephone Expenses': 'telephone_expenses',
    'Diwali Expenses (Festival Expenses)': 'festival_expenses',
    'Electricity Charges': 'electricity_charges',
    'Food & refreshment Expenses': 'food_refreshment',
    'Housekeeping Expenses': 'housekeeping_expenses',
    'Office Expenses': 'office_expenses',
    'Printing & Stationery Exp.': 'printing_Stationery',
    'Staff Welfare Expenses': 'staff_exp',
    'Stipend Expenses': 'stipend_expenses',
    'Misc. Exp': 'misc_exp',
    'Repair & maintenance': 'repair_maintenance_period',
    'travelling and conveyance Exp': 'travelling_conveyance_exp',
    'Hotel accomodation': 'hotel_accomodation',
    'Water Charges': 'water_charges',
    'Rent': 'rent',
    'Total': 'total_exp'
}

STRING_COLUMNS = [
    'region_name', 'branch_name', 'actual_branch', 'branch_location',
    'branch_manager', 'rh_name_name', 'repair_maintenance_period'
]

FLOAT_COLUMNS = [
    'courier_charges', 'internet_charges', 'telephone_expenses', 'festival_expenses',
    'electricity_charges', 'food_refreshment', 'housekeeping_expenses', 'office_expenses',
    'printing_Stationery', 'staff_exp', 'stipend_expenses', 'misc_exp',
    'travelling_conveyance_exp', 'hotel_accomodation', 'water_charges', 'rent', 'total_exp'
]

REQUIRED_COLUMNS = (
    ['region_name', 'branch_name', 'actual_branch', 'branch_location', 'branch_manager',
     'rh_name_name', 'no_of_emp', 'month_period'] + FLOAT_COLUMNS
)


def convert_month_format(month_str):
    """Convert a month string like 'Jun-25' or 'June 2025' to a proper date."""
    try:
        if pd.isna(month_str) or str(month_str).lower() == 'nan':
            return None
        month_str = str(month_str).strip()
        if '-' in month_str:
            parts = month_str.split('-')
            if len(parts) == 2:
                month_name = parts[0]
                year = '20' + parts[1] if len(parts[1]) == 2 else parts[1]
                return pd.to_datetime(f"{month_name}-{year}", format='%b-%Y', errors='coerce')
        return pd.to_datetime(month_str, errors='coerce')
    except Exception:
        return None


def convert_excel_to_csv(input_file_path, output_file_path):
    """Convert Excel file to CSV with column header mapping and type conversion."""
    try:
        print(f"Reading Excel file: {input_file_path}")
        df = pd.read_excel(input_file_path)

        print("\nOriginal columns found:")
        for i, col in enumerate(df.columns):
            print(f"{i + 1}. {col}")

        df_renamed = df.rename(columns=COLUMN_MAPPING)

        unmapped_cols = [col for col in df.columns if col not in COLUMN_MAPPING]
        if unmapped_cols:
            print(f"\nWarning: unmapped columns found: {unmapped_cols}")

        print("\nApplying data type conversions...")

        for col in STRING_COLUMNS:
            if col in df_renamed.columns:
                df_renamed[col] = df_renamed[col].astype(str).replace('nan', '').str.strip()

        if 'no_of_emp' in df_renamed.columns:
            df_renamed['no_of_emp'] = pd.to_numeric(
                df_renamed['no_of_emp'], errors='coerce'
            ).fillna(0).astype(int)

        for col in FLOAT_COLUMNS:
            if col in df_renamed.columns:
                df_renamed[col] = pd.to_numeric(df_renamed[col], errors='coerce').fillna(0.0)

        if 'month_period' in df_renamed.columns:
            df_renamed['month_period'] = df_renamed['month_period'].apply(convert_month_format)

        for col in REQUIRED_COLUMNS:
            if col not in df_renamed.columns:
                if col == 'no_of_emp':
                    df_renamed[col] = 0
                elif col == 'month_period':
                    df_renamed[col] = None
                elif col in STRING_COLUMNS:
                    df_renamed[col] = ''
                else:
                    df_renamed[col] = 0.0

        df_final = df_renamed[REQUIRED_COLUMNS]

        print(f"Saving to CSV: {output_file_path}")
        df_final.to_csv(output_file_path, index=False, encoding='utf-8')

        print("\nConversion completed successfully!")
        print(f"Total rows processed: {len(df_final)}")
        print(f"Total columns: {len(df_final.columns)}")

        return df_final

    except Exception as e:
        print(f"Error during conversion: {str(e)}")
        return None


def main():
    parser = argparse.ArgumentParser(
        description="Convert a branch expense Excel report to a SQL-ready CSV."
    )
    parser.add_argument("--input", required=True, help="Path to input Excel file")
    parser.add_argument("--output", required=True, help="Path for output CSV file")
    args = parser.parse_args()

    result_df = convert_excel_to_csv(args.input, args.output)
    if result_df is not None:
        print(f"\nOutput file saved as: {args.output}")
    else:
        print("\nConversion failed. See error messages above.")


if __name__ == "__main__":
    main()
