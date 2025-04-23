import pandas as pd
import logging
from config import INVENTORY_COLUMNS
from processing.status import get_enhanced_status, get_enhanced_recommendation

logger = logging.getLogger(__name__)

def process_data(df_master, df_responses):
    """Process the input data and generate analysis results.
    
    Args:
        df_master (DataFrame): Master data with forecast information
        df_responses (DataFrame): Plant response data with requested quantities
        
    Returns:
        dict: Dictionary containing processed data frames
    """
    logger.info("\nProcessing data...")
    
    # Get unique plants from responses
    unique_plants = sorted(df_responses['Plant'].unique())
    logger.info(f"Found {len(unique_plants)} unique plants in responses: {', '.join(unique_plants)}")
    
    # Calculate total requests by item
    logger.info("Calculating total plant requests...")
    total_requests = df_responses.groupby('Item Code')['Qty Needed'].sum().reset_index()
    total_requests.rename(columns={'Qty Needed': 'Total Plant Requests'}, inplace=True)
    
    # Calculate requests by plant
    plant_requests = {}
    for plant in unique_plants:
        plant_data = df_responses[df_responses['Plant'] == plant].groupby('Item Code')['Qty Needed'].sum().reset_index()
        plant_data.rename(columns={'Qty Needed': f'{plant} Requests'}, inplace=True)
        plant_requests[plant] = plant_data
    
    # Ensure all required columns exist in df_master
    available_columns = [col for col in INVENTORY_COLUMNS if col in df_master.columns]
    
    # Create comparison dataframe
    logger.info("Creating comparison analysis...")
    comparison = pd.merge(df_master[available_columns], total_requests, on='Item Code', how='left')
    
    # Fill NaN values for items with no requests
    comparison['Total Plant Requests'] = comparison['Total Plant Requests'].fillna(0)
    
    # Add plant-specific request columns
    for plant, plant_df in plant_requests.items():
        comparison = pd.merge(comparison, plant_df, on='Item Code', how='left')
        comparison[f'{plant} Requests'] = comparison[f'{plant} Requests'].fillna(0)
    
    # Ensure numeric columns are properly formatted - replace NaN with 0
    numeric_cols = ['Stock Qty', 'Open PRs Total 24 Months', 'Open POs Total 24 Months', 'Pr Not Confirmed 24 Months']
    for col in numeric_cols:
        if col in comparison.columns:
            comparison[col] = comparison[col].fillna(0)
    
    # NEW: Convert Annual Forecast values less than 1 to 0
    comparison['Adjusted Annual Forecast'] = comparison['Annual Forecast'].copy()
    comparison.loc[comparison['Annual Forecast'] < 1, 'Adjusted Annual Forecast'] = 0
    
    # Calculate difference using Adjusted Annual Forecast
    comparison['Difference'] = comparison['Total Plant Requests'] - comparison['Adjusted Annual Forecast']
    
    # NEW: Calculate total available inventory (stock + open orders)
    comparison['Available Stock'] = comparison['Stock Qty'] if 'Stock Qty' in comparison.columns else 0
    
    # NEW: Calculate total pending orders
    comparison['Pending Orders'] = (
        (comparison['Open PRs Total 24 Months'] if 'Open PRs Total 24 Months' in comparison.columns else 0) +
        (comparison['Open POs Total 24 Months'] if 'Open POs Total 24 Months' in comparison.columns else 0) +
        (comparison['Pr Not Confirmed 24 Months'] if 'Pr Not Confirmed 24 Months' in comparison.columns else 0)
    )
    
    # NEW: Calculate net difference after considering stock and pending orders
    comparison['Net Difference'] = comparison['Difference'] - comparison['Available Stock'] - comparison['Pending Orders']
    
    # Apply enhanced status classification based on difference and inventory
    comparison['Status'] = comparison.apply(get_enhanced_status, axis=1)
    
    # Add enhanced recommendation
    comparison['Recommendation'] = comparison.apply(get_enhanced_recommendation, axis=1)
    
    # Convert numeric columns to integers
    numeric_cols = [
        'Annual Forecast', 'Adjusted Annual Forecast', 'Total Plant Requests', 
        'Difference', 'Net Difference', 'Available Stock', 'Pending Orders'
    ]
    
    for col in numeric_cols:
        if col in comparison.columns:
            comparison[col] = comparison[col].astype(int)
    
    # Convert plant request columns to integers
    for plant in unique_plants:
        comparison[f'{plant} Requests'] = comparison[f'{plant} Requests'].astype(int)
    
    # Generate summary statistics
    logger.info("Generating summary statistics...")
    summary_stats = calculate_summary_statistics(comparison)
    
    # Generate plant summary
    plant_summary = calculate_plant_summary(comparison, unique_plants)
    
    return {
        'comparison': comparison,
        'summary_stats': summary_stats,
        'plant_summary': plant_summary,
        'unique_plants': unique_plants
    }


def calculate_summary_statistics(comparison):
    """Generate enhanced summary statistics from comparison data.
    
    Args:
        comparison (DataFrame): Processed comparison data
        
    Returns:
        DataFrame: Summary statistics
    """
    
    summary_stats = pd.DataFrame({
        'Metric': [
            'Total Items Analyzed',
            'Items with High Deviation (>3 after stock & orders)',
            'Items with Moderate Deviation (1-3 after stock & orders)',
            'Items with Acceptable Match (=0)',
            'Items with Low Request (<0)',
            'Items Covered by Available Stock',
            'Items Covered by Pending Orders',
            'Items Needing Significant Increase',
            'Items Needing Moderate Increase',
            'Items Needing Forecast Reduction',
            'Items with Adequate Forecast',
            'Items with Forecast < 1 (Adjusted to 0)',
            'Items with Stock Available',
            'Items with Pending Orders'
        ],
        'Count': [
            len(comparison),
            len(comparison[comparison['Status'] == 'HIGH_DEVIATION']),
            len(comparison[comparison['Status'] == 'MODERATE_DEVIATION']),
            len(comparison[comparison['Status'] == 'ACCEPTABLE']),
            len(comparison[comparison['Status'] == 'LOW_REQUEST']),
            len(comparison[comparison['Status'] == 'COVERED_BY_STOCK']),
            len(comparison[comparison['Status'] == 'COVERED_BY_ORDERS']),
            len(comparison[comparison['Recommendation'] == 'Significant increase needed after using stock & orders']),
            len(comparison[comparison['Recommendation'] == 'Moderate increase needed after using stock & orders']),
            len(comparison[comparison['Recommendation'] == 'Consider reducing forecast']),
            len(comparison[comparison['Recommendation'] == 'Current forecast appears adequate']),
            len(comparison[comparison['Annual Forecast'] < 1]),
            len(comparison[comparison['Available Stock'] > 0]),
            len(comparison[comparison['Pending Orders'] > 0])
        ]
    })
    
    return summary_stats


def calculate_plant_summary(comparison, unique_plants):
    """Generate plant summary statistics.
    
    Args:
        comparison (DataFrame): Processed comparison data
        unique_plants (list): List of unique plant names
        
    Returns:
        DataFrame: Plant summary statistics
    """
    
    plant_data = []
    
    for plant in unique_plants:
        # Get items requested by this plant (where requests > 0)
        plant_items = comparison[comparison[f'{plant} Requests'] > 0]
        
        # Count by status
        high_dev = len(plant_items[plant_items['Status'] == 'HIGH_DEVIATION'])
        moderate = len(plant_items[plant_items['Status'] == 'MODERATE_DEVIATION'])
        acceptable = len(plant_items[plant_items['Status'] == 'ACCEPTABLE'])
        low_req = len(plant_items[plant_items['Status'] == 'LOW_REQUEST'])
        covered_by_stock = len(plant_items[plant_items['Status'] == 'COVERED_BY_STOCK'])
        covered_by_orders = len(plant_items[plant_items['Status'] == 'COVERED_BY_ORDERS'])
        
        plant_data.append({
            'Plant': plant,
            'Items Requested': len(plant_items),
            'High Deviation Items': high_dev,
            'Moderate Deviation Items': moderate,
            'Acceptable Items': acceptable,
            'Low Request Items': low_req,
            'Covered by Stock': covered_by_stock,
            'Covered by Orders': covered_by_orders
        })
    
    return pd.DataFrame(plant_data)