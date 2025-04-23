import logging
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
import pandas as pd
from config import OUTPUT_PATH
from processing.status import get_enhanced_status_for_plant, get_enhanced_recommendation_for_plant
from visualization.formatters.comparison import format_comparison_sheet
from visualization.formatters.plant import format_plant_sheet
from visualization.formatters.instructions import create_instructions_sheet
from visualization.formatters.dashboard import create_dashboard

logger = logging.getLogger(__name__)

def create_output_file(comparison, plant_summary, summary_stats, unique_plants, df_master, df_responses):
    """Create and format the output Excel file.
    
    Args:
        comparison (DataFrame): Processed comparison data
        plant_summary (DataFrame): Plant summary statistics
        summary_stats (DataFrame): Overall summary statistics
        unique_plants (list): List of unique plant names
        df_master (DataFrame): Original master data
        df_responses (DataFrame): Original response data
    """
    logger.info("\nCreating output file...")
    
    # Create Excel writer
    writer = pd.ExcelWriter(OUTPUT_PATH, engine='openpyxl')
    
    # Write source data
    df_master.to_excel(writer, sheet_name='Master', index=False)
    df_responses.to_excel(writer, sheet_name='Responses', index=False)
    
    # Write analysis results
    comparison.to_excel(writer, sheet_name='Comparison Analysis', index=False)
    summary_stats.to_excel(writer, sheet_name='Summary Statistics', index=False)
    plant_summary.to_excel(writer, sheet_name='Plant Summary', index=False)
    
    # Create plant-specific sheets
    for plant in unique_plants:
        logger.info(f"Creating plant sheet for {plant}...")
        create_plant_sheet(writer, plant, df_master, df_responses, comparison)
    
    # Save the workbook to access it with openpyxl
    writer.close()
    
    # Open the file with openpyxl to add formatting and charts
    wb = openpyxl.load_workbook(OUTPUT_PATH)
    
    # Add instructions sheet
    create_instructions_sheet(wb)
    
    # Create dashboard
    create_dashboard(wb, comparison, summary_stats, plant_summary, unique_plants)
    
    # Format the comparison sheet
    format_comparison_sheet(wb, comparison)
    
    # Format plant sheets
    for plant in unique_plants:
        # Get shortened sheet name
        short_plant_name = plant[:25] if len(plant) > 25 else plant
        sheet_name = f'Plant_{short_plant_name}'
        
        if sheet_name in wb.sheetnames:
            format_plant_sheet(wb, plant, sheet_name)
        else:
            logger.warning(f"Sheet '{sheet_name}' not found")
    
    # Save the workbook
    wb.save(OUTPUT_PATH)


def create_plant_sheet(writer, plant, df_master, df_responses, comparison):
    """Create a sheet for plant-specific data and communication.
    
    Args:
        writer: Excel writer object
        plant (str): Plant name
        df_master (DataFrame): Master data
        df_responses (DataFrame): Response data
        comparison (DataFrame): Comparison analysis data
    """
    # Truncate plant name to fit Excel's 31-character limit for sheet names
    short_plant_name = plant[:25] if len(plant) > 25 else plant
    sheet_name = f'Plant_{short_plant_name}'
    
    # Filter responses for this plant
    plant_responses = df_responses[df_responses['Plant'] == plant].copy()
    
    # Get unique item codes requested by this plant
    plant_item_codes = plant_responses['Item Code'].unique()
    
    # Filter master data for these item codes and get required columns
    master_columns = [
        'Item Code', 'Description', 'Annual Forecast', 
        'Accumulated (36m)', 'Accumulated (24m)', 'Accumulated (12m)',
        'Stock Qty', 'Open PRs Total 24 Months', 'Open POs Total 24 Months',
        'Pr Not Confirmed 24 Months', 'Code Creation Date', 
        'Classification Calculated', 'Projects'
    ]
    
    # Ensure all required columns exist in df_master
    available_columns = [col for col in master_columns if col in df_master.columns]
    
    plant_master = df_master[df_master['Item Code'].isin(plant_item_codes)][available_columns].copy()
    
    # Get plant-specific comparison data
    plant_comparison = comparison[comparison[f'{plant} Requests'] > 0].copy()
    
    # Calculate plant-specific difference and status
    plant_comparison['Plant Difference'] = plant_comparison[f'{plant} Requests'] - plant_comparison['Adjusted Annual Forecast']
    
    # Calculate plant-specific net difference
    plant_comparison['Plant Net Difference'] = plant_comparison['Plant Difference'] - \
                                               plant_comparison['Available Stock'] - \
                                               plant_comparison['Pending Orders']
    
    # Apply status classification
    plant_comparison['Plant Status'] = plant_comparison.apply(
        lambda row: get_enhanced_status_for_plant(row, plant), 
        axis=1
    )
    
    plant_comparison['Plant Recommendation'] = plant_comparison.apply(
        lambda row: get_enhanced_recommendation_for_plant(row, plant),
        axis=1
    )
    
    # Extract relevant columns for the plant sheet
    plant_data_columns = [
        'Item Code', f'{plant} Requests', 'Plant Difference', 
        'Available Stock', 'Pending Orders', 'Plant Net Difference',
        'Plant Status', 'Plant Recommendation'
    ]
    plant_specific_data = plant_comparison[plant_data_columns].copy()
    
    # Combine the data
    # First, prepare a merged dataframe with all master data and plant-specific comparison
    plant_master_comparison = pd.merge(
        plant_master, 
        plant_specific_data, 
        on='Item Code', 
        how='left'
    )
    
    # Then merge with the responses to get all columns from responses
    plant_responses_with_comparison = pd.merge(
        plant_responses,
        plant_master_comparison,
        on='Item Code',
        how='left',
        suffixes=('', '_master')
    )
    
    # Remove any duplicate columns
    duplicate_cols = [col for col in plant_responses_with_comparison.columns if col.endswith('_master')]
    plant_responses_with_comparison = plant_responses_with_comparison.drop(columns=duplicate_cols)
    
    # Write to Excel with the shortened sheet name
    plant_responses_with_comparison.to_excel(writer, sheet_name=sheet_name, index=False)