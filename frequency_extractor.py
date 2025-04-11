#!/usr/bin/env python
# -*- coding: utf-8 -*-

import pandas as pd
import numpy as np
from pathlib import Path
from datetime import datetime
import glob
import os

def extract_specific_frequencies(input_file: str, target_frequencies: list = None) -> pd.DataFrame:
    """Extract specific frequency points from an Excel file.
    
    Args:
        input_file (str): Path to the input Excel file.
        target_frequencies (list, optional): List of frequencies to extract.
            If None, uses frequencies from 50 to 22000 Hz in steps of 50.
            Defaults to None.
    
    Returns:
        pd.DataFrame: DataFrame containing the extracted frequency points and their magnitude values.
    
    Example:
        >>> df = extract_specific_frequencies("HL1/Pon.xls")
        >>> print(df.head())
    """
    # Read Excel file
    df = pd.read_excel(input_file)
    
    # Try to identify the frequency and magnitude columns
    freq_column = None
    mag_column = None
    
    # Define possible column name patterns
    freq_patterns = ['Frequency', 'frequency', 'freq', 'Freq', 'Hz', 'hz']
    mag_patterns = ['dBSPL', 'dbspl', 'mag', 'magnitude', 'DB', 'db']
    
    # Find frequency column
    for col in df.columns:
        if any(pattern in str(col).lower() for pattern in freq_patterns):
            freq_column = col
            break
    
    # Find magnitude column
    for col in df.columns:
        if any(pattern in str(col).lower() for pattern in mag_patterns):
            mag_column = col
            break
    
    if freq_column is None:
        # If no frequency column found, assume first column is frequency
        freq_column = df.columns[0]
        print(f"Assuming first column '{freq_column}' contains frequencies")
    
    if mag_column is None:
        # If no magnitude column found, assume second column is magnitude
        mag_column = df.columns[1]
        print(f"Assuming second column '{mag_column}' contains magnitude values")
    
    # If no specific frequencies provided, generate range from 50 to 22000 in steps of 50
    if target_frequencies is None:
        target_frequencies = list(range(50, 22050, 50))
    
    # Find closest frequency points in the data
    result_rows = []
    for target_freq in target_frequencies:
        # Find the closest frequency in the dataset
        closest_idx = (df[freq_column] - target_freq).abs().idxmin()
        # Only keep frequency and magnitude columns
        result_row = pd.Series({
            freq_column: df.loc[closest_idx, freq_column],
            mag_column: df.loc[closest_idx, mag_column]
        })
        result_rows.append(result_row)
    
    result_df = pd.DataFrame(result_rows)
    
    # Get HL folder name and file name
    path = Path(input_file)
    hl_folder = path.parent.name
    file_name = path.stem
    date_str = datetime.now().strftime('%Y%m%d')
    exp_num = Path.cwd().name
    
    # Generate output filename including HL folder name
    output_file = f"{exp_num}_{date_str}_{hl_folder}_{file_name}_magnitude.xlsx"
    output_path = Path("extracted_frequencies") / output_file
    
    # Create output directory if it doesn't exist
    output_path.parent.mkdir(exist_ok=True)
    
    # Save results
    result_df.to_excel(output_path, index=False)
    print(f"Results saved to {output_path}")
    
    return result_df

def process_folder(folder_path: str, target_frequencies: list = None) -> None:
    """Process all .xls files in the specified folder.
    
    Args:
        folder_path (str): Path to the folder containing .xls files.
        target_frequencies (list, optional): List of specific frequencies to extract.
            If None, will use frequencies from 50 to 22000 Hz in steps of 50.
            Defaults to None.
    """
    xls_files = glob.glob(os.path.join(folder_path, "*.xls"))
    
    for file_path in xls_files:
        try:
            print(f"\nProcessing {file_path}...")
            result = extract_specific_frequencies(file_path, target_frequencies)
            print(f"Successfully processed {file_path}")
        except Exception as e:
            print(f"Error processing {file_path}: {str(e)}")

def main():
    """Main function to process the frequency points extraction for all HL folders."""
    # Generate frequencies from 50 to 22000 in steps of 50
    target_frequencies = list(range(50, 22050, 50))
    
    # Get all HL folders
    hl_folders = glob.glob("HL*")
    
    if not hl_folders:
        print("No HL folders found!")
        return
    
    print(f"Found {len(hl_folders)} HL folders")
    
    # Process each folder
    for folder in sorted(hl_folders):
        print(f"\nProcessing folder {folder}...")
        process_folder(folder, target_frequencies)
        
    print("\nAll processing completed!")

if __name__ == "__main__":
    main()