"""
policy_data_sql_prep.py

Prepares raw policy/transaction Excel exports for SQL loading:
  - normalizes numeric ID columns (c_no, inst_no) that Excel/pandas otherwise
    mangle into floats (e.g. 100234.0 -> "100234")
  - builds a composite primary key from those normalized IDs
  - parses date columns and computes retention (payin - payout)
  - adds a '~' prefix to ID/date columns as a load-safety marker some SQL
    bulk-import tools expect, with a companion cleaning step to strip it
    back out and normalize a few known placeholder values afterwards

Run interactively; it will prompt for one or more input Excel files (to
process into CSV) or an already-processed CSV (to clean).

Usage:
    python policy_data_sql_prep.py
"""

import os
import traceback
from datetime import date

import numpy as np
import pandas as pd


def process_financial_data(file_path):
    """Process a single raw Excel export into a SQL-load-ready DataFrame."""
    try:
        df = pd.read_excel(file_path)

        df['payin_amount'] = df['payin_amount'].astype(float)
        df['payout_amount'] = df['payout_amount'].astype(float)

        date_columns = ['start_date', 'expiry_date', 'tp_expiry_date']
        for col in date_columns:
            df[col] = pd.to_datetime(df[col], errors='coerce')

        # c_no is numeric but pandas often loads it as float (e.g. 100234.0) -
        # normalize back to a clean string before it becomes part of the key.
        if 'c_no' in df.columns:
            df['c_no'] = df['c_no'].fillna("")

            def clean_c_no(x):
                if x == "" or pd.isna(x) or str(x).lower() == 'nan':
                    return ""
                try:
                    num_val = float(x)
                    return str(int(num_val)) if num_val.is_integer() else str(num_val)
                except Exception:
                    return str(x)

            df['c_no'] = df['c_no'].apply(clean_c_no)

        if 'inst_no' in df.columns:
            df['inst_no'] = df['inst_no'].fillna("").astype(str).replace('nan', '')

        # Composite primary key (built after c_no/inst_no are normalized)
        df['pk'] = df['c_no'] + df['inst_no']

        columns_to_clean = ['pk', 'policy_no', 'start_date', 'expiry_date', 'tp_expiry_date']
        df[columns_to_clean] = df[columns_to_clean].fillna("")

        # '~' prefix marks these columns for the downstream SQL loader
        prefix_columns = ['c_no', 'inst_no', 'pk', 'policy_no', 'start_date', 'expiry_date', 'tp_expiry_date']
        for col in prefix_columns:
            if col not in df.columns:
                continue
            if col in ['c_no', 'inst_no', 'pk']:
                df[col] = df[col].apply(lambda x: '' if str(x) == '' or str(x) == 'nan' else '~' + str(x))
            else:
                df[col] = df[col].astype(str)
                df[col] = df[col].apply(lambda x: '' if x in ('', 'NaT') else '~' + x)
                df[col] = df[col].replace('~NaT', '')

        df['retention'] = df['payin_amount'] - df['payout_amount']
        df['nop'] = df['nop'].apply(lambda x: 0 if pd.notna(x) and x != 0 else x)
        df['entry_date_db'] = date.today().strftime("%Y-%m-%d")

        return df
    except Exception:
        print(f"Error processing {file_path}:")
        print(traceback.format_exc())
        return None


def clean_csv_data(csv_path):
    """
    Strip the '~' load-safety prefix back out of a processed CSV, normalize
    a couple of known placeholder values, and save alongside the original.
    """
    try:
        df = pd.read_csv(csv_path)

        columns_to_clean = ['c_no', 'inst_no', 'pk', 'start_date', 'expiry_date', 'tp_expiry_date']
        for col in columns_to_clean:
            if col in df.columns:
                df[col] = df[col].astype(str).str.replace('~', '')

        code_columns = ['rm_code', 'ref_pos_code']
        for col in code_columns:
            if col in df.columns:
                df[col] = df[col].astype(str).replace('DIRECT', '0')

        if 'tp_expiry_date' in df.columns:
            df['tp_expiry_date'] = df['tp_expiry_date'].replace(['', 'blank', 'empty'], np.nan)

        output_path = csv_path.replace('.csv', '_cleaned.csv')
        df.to_csv(output_path, index=False)

        print("\n--- Cleaning Summary ---")
        print(f"Input file: {csv_path}")
        print(f"Output file: {output_path}")
        print(f"Total records processed: {len(df)}")

        return df
    except Exception as e:
        print(f"Error processing {csv_path}: {e}")
        return None


def main():
    processed_output_dir = 'financial_processing_output'
    cleaned_output_dir = 'financial_cleaning_output'
    os.makedirs(processed_output_dir, exist_ok=True)
    os.makedirs(cleaned_output_dir, exist_ok=True)

    while True:
        print("\n--- FINANCIAL DATA PROCESSING TOOL ---")
        print("1. Process Excel Files")
        print("2. Clean Processed CSV Files")
        print("3. Exit")

        mode_choice = input("Enter your choice (1/2/3): ").strip()

        if mode_choice == '1':
            while True:
                try:
                    num_files = int(input("Enter number of Excel files to process: "))
                    if num_files > 0:
                        break
                    print("Please enter a positive number.")
                except ValueError:
                    print("Invalid input. Enter a number.")

            file_paths = []
            print("\nEnter FULL file paths:")
            for i in range(num_files):
                while True:
                    file_path = input(f"File {i + 1} path: ").strip().replace('"', '')
                    if os.path.exists(file_path) and file_path.lower().endswith(('.xlsx', '.xls')):
                        file_paths.append(file_path)
                        break
                    print("Invalid path. Ensure file exists and is an Excel file.")

            successful_files, failed_files = 0, 0
            print("\n--- PROCESSING STARTED ---")
            for idx, file_path in enumerate(file_paths, 1):
                try:
                    print(f"\nProcessing File {idx}: {os.path.basename(file_path)}")
                    processed_df = process_financial_data(file_path)
                    if processed_df is not None:
                        filename = os.path.splitext(os.path.basename(file_path))[0]
                        csv_path = os.path.join(processed_output_dir, f'cleaned_{filename}.csv')
                        processed_df.to_csv(csv_path, index=False)

                        print(f"\n--- Summary for {filename} ---")
                        print(f"Total Records: {len(processed_df)}")
                        print(f"Total Payin Amount: {processed_df['payin_amount'].sum():,.2f}")
                        print(f"Total Payout Amount: {processed_df['payout_amount'].sum():,.2f}")
                        print(f"Total Gross Premium: {processed_df['gross_premium'].sum():,.2f}")
                        print(f"Total Retention: {processed_df['retention'].sum():,.2f}")
                        print(f"CSV saved to: {csv_path}")
                        successful_files += 1
                    else:
                        failed_files += 1
                except Exception:
                    print(f"Unexpected error processing {file_path}:")
                    print(traceback.format_exc())
                    failed_files += 1

            print("\n--- PROCESSING COMPLETE ---")
            print(f"Total Files Processed: {num_files}")
            print(f"Successful Files: {successful_files}")
            print(f"Failed Files: {failed_files}")

        elif mode_choice == '2':
            while True:
                csv_path = input("Enter the path to your CSV file ('b' to go back, 'q' to quit): ").strip().replace('"', '')
                if csv_path.lower() == 'q':
                    return
                if csv_path.lower() == 'b':
                    break
                if not csv_path.endswith('.csv'):
                    print("Please provide a valid CSV file path")
                    continue

                cleaned_df = clean_csv_data(csv_path)
                print("\nCleaning completed successfully!" if cleaned_df is not None else "\nFailed to clean the CSV file")
                print("\n" + "=" * 50 + "\n")

        elif mode_choice == '3':
            print("Exiting.")
            break
        else:
            print("Invalid choice. Please enter 1, 2, or 3.")


if __name__ == "__main__":
    main()
