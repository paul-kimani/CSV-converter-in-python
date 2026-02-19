"""
Simple Excel to CSV Converter
- Converts all Excel files (.xls, .xlsx, etc.) to CSV
- Extracts ONLY the first sheet from each file
- Preserves original folder structure
- Robust error handling for corrupt files
- Saves with same name as original (just .csv extension)
"""

import os
import pandas as pd
from pathlib import Path
import logging
from datetime import datetime

# ==================== CONFIGURATION ====================
SOURCE_FOLDER = "source folder path here"  # <-- Set this to your source folder containing Excel files
OUTPUT_FOLDER = "output folder path here"  # <-- Set this to your desired output folder for CSV files
LOG_FOLDER = "log folder path here"  # <-- Set this to your desired log folder for logs and failed file records

# ==================== SETUP ====================
os.makedirs(OUTPUT_FOLDER, exist_ok=True)
os.makedirs(LOG_FOLDER, exist_ok=True)

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(message)s',
    handlers=[
        logging.FileHandler(os.path.join(LOG_FOLDER, 'conversion.log')),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# ==================== MAIN CONVERSION FUNCTION ====================
def convert_excel_to_csv():
    """Convert all Excel files to CSV, extracting only the first sheet"""
    
    # Statistics
    total = 0
    succeeded = 0
    failed = 0
    skipped = 0
    
    # Find all Excel files
    excel_extensions = ['.xls', '.xlsx', '.xlsm', '.xlsb']
    excel_files = []
    
    for ext in excel_extensions:
        excel_files.extend(Path(SOURCE_FOLDER).rglob(f'*{ext}'))
    
    logger.info(f"Found {len(excel_files)} Excel files to process")
    
    # Process each file
    for file_path in excel_files:
        total += 1
        
        try:
            # Create output path maintaining folder structure
            relative_path = file_path.relative_to(SOURCE_FOLDER)
            output_path = Path(OUTPUT_FOLDER) / relative_path.parent / (file_path.stem + '.csv')
            output_path.parent.mkdir(parents=True, exist_ok=True)
            
            # Skip if CSV already exists
            if output_path.exists():
                logger.info(f"SKIP: {file_path.name} -> already exists")
                skipped += 1
                continue
            
            logger.info(f"Processing: {file_path.name}")
            
            # Try different engines to read the Excel file
            df = None
            engines_to_try = ['xlrd', 'openpyxl', None]  # None lets pandas decide
            
            for engine in engines_to_try:
                try:
                    if engine:
                        df = pd.read_excel(file_path, sheet_name=0, engine=engine)
                    else:
                        df = pd.read_excel(file_path, sheet_name=0)
                    
                    if df is not None and not df.empty:
                        logger.info(f"  → Read successfully with engine: {engine or 'default'}")
                        break
                        
                except Exception as e:
                    continue
            
            # If all engines failed, try one more time with different parameters
            if df is None:
                try:
                    # Try reading without specifying engine, with header detection off
                    df = pd.read_excel(file_path, sheet_name=0, header=None)
                except:
                    pass
            
            # If we got data, save it
            if df is not None and not df.empty:
                df.to_csv(output_path, index=False)
                logger.info(f"  → Saved: {output_path.name} ({len(df)} rows)")
                succeeded += 1
            else:
                raise Exception("No data could be read from file")
                
        except Exception as e:
            logger.error(f"FAIL: {file_path.name} - {str(e)}")
            failed += 1
            
            # Log failed files
            with open(os.path.join(LOG_FOLDER, 'failed_files.txt'), 'a') as f:
                f.write(f"{datetime.now()}\t{file_path}\t{str(e)}\n")
    
    # Print summary
    logger.info("=" * 50)
    logger.info("CONVERSION COMPLETE")
    logger.info("=" * 50)
    logger.info(f"Total files: {total}")
    logger.info(f"Successfully converted: {succeeded}")
    logger.info(f"Failed: {failed}")
    logger.info(f"Skipped (already exist): {skipped}")
    logger.info("=" * 50)
    logger.info(f"Output folder: {OUTPUT_FOLDER}")
    logger.info(f"Log folder: {LOG_FOLDER}")

# ==================== RUN ====================
if __name__ == "__main__":
    convert_excel_to_csv()