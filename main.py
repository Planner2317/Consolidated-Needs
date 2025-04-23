import pandas as pd
import os
import logging
from utils.logging_setup import setup_logging
from config import MASTER_PATH, RESPONSES_PATH, OUTPUT_PATH
from processing.data_processor import process_data
from visualization.excel_output import create_output_file

# Get logger
logger = setup_logging()

def main():
    """Main function to run the gasket inventory analysis."""
    logger.info("Starting Gasket Inventory Analysis...")
    
    try:
        # Load source data files
        logger.info("\nReading master file...")
        df_master = pd.read_excel(MASTER_PATH)
        logger.info(f"Successfully loaded {len(df_master)} items from Master file")
        
        logger.info("Reading responses file...")
        df_responses = pd.read_excel(RESPONSES_PATH)
        logger.info(f"Successfully loaded {len(df_responses)} response records")
        
        # Process the data
        result_data = process_data(df_master, df_responses)
        
        # Generate output file
        create_output_file(
            result_data['comparison'], 
            result_data['plant_summary'], 
            result_data['summary_stats'], 
            result_data['unique_plants'],
            df_master, 
            df_responses
        )
        
        logger.info(f"\nAnalysis complete! Output file saved to: {OUTPUT_PATH}")
        
    except FileNotFoundError as e:
        logger.error(f"File not found error: {str(e)}")
    except pd.errors.EmptyDataError:
        logger.error("One of the Excel files is empty or has no valid data")
    except pd.errors.ParserError:
        logger.error("Error parsing Excel file - file may be corrupted")
    except Exception as e:
        logger.error(f"Unexpected error: {str(e)}")

if __name__ == "__main__":
    main()