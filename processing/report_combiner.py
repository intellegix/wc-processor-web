"""
Workers Compensation Report Combiner
Combines two workers comp CSV reports for companies on the same policy.
"""

import pandas as pd
from datetime import datetime
import os


def read_file_smart(file_path):
    """
    Read file intelligently - handles both Excel (.xlsx) and CSV files.
    Automatically detects file type and encoding.
    """
    file_ext = os.path.splitext(file_path)[1].lower()

    # Handle Excel files
    if file_ext in ['.xlsx', '.xls']:
        try:
            df = pd.read_excel(file_path, sheet_name=0)
            print(f"  Successfully read Excel file")
            return df
        except Exception as e:
            raise ValueError(f"Unable to read Excel file: {e}")

    # Handle CSV files with encoding detection
    encodings = ['utf-8', 'latin-1', 'cp1252', 'iso-8859-1', 'utf-16']

    for encoding in encodings:
        try:
            df = pd.read_csv(file_path, encoding=encoding)
            print(f"  Successfully read CSV with {encoding} encoding")
            return df
        except (UnicodeDecodeError, UnicodeError):
            continue
        except Exception:
            continue

    # If all encodings fail, try with error handling
    try:
        df = pd.read_csv(file_path, encoding='utf-8', errors='ignore')
        print(f"  Read CSV with UTF-8 (ignoring errors)")
        return df
    except Exception as e:
        raise ValueError(f"Unable to read file with any method: {e}")


def combine_reports(file1_path, file2_path, output_dir):
    """
    Combine two workers comp reports into a single CSV file.

    Args:
        file1_path: Path to the first CSV file (ArmorPro report)
        file2_path: Path to the second CSV file (ASR report)
        output_dir: Directory where the combined report will be saved

    Returns:
        str: Path to the combined output file
    """
    print(f"Reading {os.path.basename(file1_path)}...")
    df1 = read_file_smart(file1_path)

    print(f"Reading {os.path.basename(file2_path)}...")
    df2 = read_file_smart(file2_path)

    # Display summary information
    print(f"\nFirst Report: {len(df1)} rows")
    print(f"Second Report: {len(df2)} rows")

    # Combine the dataframes
    print("\nCombining reports...")
    combined_df = pd.concat([df1, df2], ignore_index=True)

    # Generate output filename with timestamp
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_filename = f"{timestamp}_CombinedWorkersCompReport.csv"
    output_path = os.path.join(output_dir, output_filename)

    # Save the combined report
    print(f"\nSaving combined report to: {output_filename}")
    combined_df.to_csv(output_path, index=False)

    print(f"\nCombined Report: {len(combined_df)} total rows")
    print(f"Successfully created: {output_path}")

    return output_path
