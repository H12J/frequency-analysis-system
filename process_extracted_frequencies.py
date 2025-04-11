#!/usr/bin/env python
# -*- coding: utf-8 -*-

import pandas as pd
import numpy as np
from pathlib import Path
from datetime import datetime
import glob
import os
from tqdm import tqdm

def classify_frequency(freq: float, base: int = 50) -> int:
    """Classify a frequency value to the nearest base frequency multiple.
    
    Args:
        freq (float): The frequency value to classify.
        base (int, optional): Base frequency for classification. 
            Defaults to 50.
    
    Returns:
        int: The classified frequency value.
    
    Example:
        >>> classify_frequency(123.45, 50)
        100
    """
    return round(freq / base) * base

def process_extracted_file(input_file: str, base: int = 50) -> pd.DataFrame:
    """Process extracted frequency data from Excel file and classify frequencies.
    
    Args:
        input_file (str): Path to the input Excel file.
        base (int, optional): Base frequency for classification. 
            Defaults to 50.
    
    Returns:
        pd.DataFrame: DataFrame containing the classified frequencies and their 
            magnitude values.
    
    Example:
        >>> df = process_extracted_file("extracted_frequencies/HL1_Pon_magnitude_only.xlsx")
        >>> print(df.head())
    """
    # Read Excel file
    df = pd.read_excel(input_file)
    
    # Identify frequency column (first column)
    freq_column = df.columns[0]
    mag_column = df.columns[1]
    
    # Create result DataFrame with classified frequencies
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
    # Extract HL folder name from input filename
    file_parts = path.stem.split('_')
    hl_folder = next((part for part in file_parts if part.startswith('HL')), '')
    file_name = path.stem
    date_str = datetime.now().strftime('%Y%m%d')
    exp_num = Path.cwd().name
    
    output_file = f"{exp_num}_{date_str}_{hl_folder}_{file_name}_reclass{base}.xlsx"
    output_path = Path("reclassified_frequencies") / output_file
    
    # Create output directory if it doesn't exist
    output_path.parent.mkdir(exist_ok=True)
    
    # Save both detailed and grouped results
    with pd.ExcelWriter(output_path) as writer:
        result_df.to_excel(writer, sheet_name='Detailed_Results', index=False)
        grouped_df.to_excel(writer, sheet_name='Grouped_Results', index=False)
    
    print(f"Results saved to {output_path}")
    
    return grouped_df

def process_all_files(input_dir: str = "extracted_frequencies", pattern: str = "*_magnitude.xlsx", base: int = 50) -> None:
    """Process all extracted frequency files in the specified directory.
    
    Args:
        input_dir (str, optional): Directory containing the files to process. 
            Defaults to "extracted_frequencies".
        pattern (str, optional): File pattern to match. 
            Defaults to "*_magnitude.xlsx".
        base (int, optional): Base frequency for classification. 
            Defaults to 50.
            
    Example:
        >>> process_all_files()
    """
    # Get all matching files
    files = glob.glob(os.path.join(input_dir, pattern))
    
    if not files:
        print(f"No files matching pattern '{pattern}' found in '{input_dir}'")
        return
    
    print(f"Found {len(files)} files to process")
    
    # Process each file
    for file_path in tqdm(files, desc="Processing files"):
        try:
            process_extracted_file(file_path, base)
        except Exception as e:
            print(f"Error processing {file_path}: {str(e)}")

def main():
    """Main function to process frequency classification for extracted files."""
    base_freq = 50  # Base frequency for classification
    
    print(f"Starting processing with {base_freq}Hz as base frequency")
    process_all_files(base=base_freq)
    print("\nAll processing completed!")

if __name__ == "__main__":
    main()