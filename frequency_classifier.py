#!/usr/bin/env python
# -*- coding: utf-8 -*-

import pandas as pd
import numpy as np
from pathlib import Path
from datetime import datetime
import glob
import os

def classify_frequency(freq: float, base: int = 50) -> int:
    """Classify a frequency value to the nearest base frequency multiple.
    
    Args:
        freq (float): The frequency value to classify.
        base (int, optional): The base frequency for classification. 
            Defaults to 50.
    
    Returns:
        int: The classified frequency value.
    
    Example:
        >>> classify_frequency(123.45, 50)
        100
    """
    return round(freq / base) * base

def process_frequency_data(input_file: str, base: int = 50) -> pd.DataFrame:
    """Process frequency data from Excel file and classify frequencies.
    
    Args:
        input_file (str): Path to the input Excel file.
        base (int, optional): Base frequency for classification. 
            Defaults to 50.
    
    Returns:
        pd.DataFrame: DataFrame containing original and classified frequencies 
            with their magnitude values.
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
        freq_column = df.columns[0]
        print(f"Assuming first column '{freq_column}' contains frequencies")
    
    if mag_column is None:
        mag_column = df.columns[1]
        print(f"Assuming second column '{mag_column}' contains magnitude values")
    
    # Create result DataFrame with original and classified frequencies
    result_df = pd.DataFrame({
        'Original_Frequency': df[freq_column],
        'Classified_Frequency': df[freq_column].apply(lambda x: classify_frequency(x, base)),
        'Magnitude': df[mag_column]
    })
    
    # Group by classified frequency and calculate mean magnitude
    grouped_df = result_df.groupby('Classified_Frequency')['Magnitude'].agg([
        ('Mean_Magnitude', 'mean'),
        ('Count', 'count')
    ]).reset_index()
    
    # Generate output filename with HL folder name
    path = Path(input_file)
    hl_folder = path.parent.name
    file_name = path.stem
    date_str = datetime.now().strftime('%Y%m%d')
    exp_num = Path.cwd().name
    output_file = f"{exp_num}_{date_str}_{hl_folder}_{file_name}_freq{base}.xlsx"
    output_path = Path("classified_frequencies") / output_file
    
    # Create output directory if it doesn't exist
    output_path.parent.mkdir(exist_ok=True)
    
    # Save both detailed and grouped results
    with pd.ExcelWriter(output_path) as writer:
        result_df.to_excel(writer, sheet_name='Detailed_Results', index=False)
        grouped_df.to_excel(writer, sheet_name='Grouped_Results', index=False)
    
    print(f"Results saved to {output_path}")
    print(f"\nFound {len(grouped_df)} unique {base}Hz frequency groups")
    print(f"Frequency range: {grouped_df['Classified_Frequency'].min():.0f}Hz - {grouped_df['Classified_Frequency'].max():.0f}Hz")
    
    return grouped_df

def process_folder(folder_path: str, base: int = 50) -> None:
    """Process all Excel files in the specified folder.
    
    Args:
        folder_path (str): Path to the folder containing Excel files.
        base (int, optional): Base frequency for classification. 
            Defaults to 50.
    """
    xls_files = glob.glob(os.path.join(folder_path, "*.xls"))
    
    for file_path in xls_files:
        try:
            print(f"\nProcessing {file_path}...")
            result = process_frequency_data(file_path, base)
            print(f"Successfully processed {file_path}")
            print("\nFirst few grouped frequencies:")
            print(result.head())
        except Exception as e:
            print(f"Error processing {file_path}: {str(e)}")

def main():
    """Main function to process frequency classification for all HL folders."""
    # Get all HL folders
    hl_folders = glob.glob("HL*")
    base_freq = 50  # Base frequency for classification
    
    if not hl_folders:
        print("No HL folders found!")
        return
    
    print(f"Found {len(hl_folders)} HL folders")
    print(f"Using {base_freq}Hz as base frequency for classification")
    
    # Process each folder
    for folder in sorted(hl_folders):
        print(f"\nProcessing folder {folder}...")
        process_folder(folder, base_freq)
    
    print("\nAll processing completed!")

if __name__ == "__main__":
    main()