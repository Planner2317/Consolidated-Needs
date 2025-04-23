def get_enhanced_status(row):
    """Determine enhanced status based on difference, stock, and pending orders.
    
    Args:
        row: DataFrame row with difference, stock, and order information
        
    Returns:
        str: Status category
    """
    difference = row['Difference']
    available_stock = row['Available Stock']
    pending_orders = row['Pending Orders']
    net_difference = row['Net Difference']
    
    # Check if deviation is covered by available stock
    if difference > 0 and difference <= available_stock:
        return "COVERED_BY_STOCK"
    
    # Check if deviation is covered by pending orders
    elif difference > 0 and difference <= (available_stock + pending_orders):
        return "COVERED_BY_ORDERS"
    
    # Otherwise use the standard classification
    elif net_difference == 0:
        return "ACCEPTABLE"
    elif 1 <= net_difference <= 3:
        return "MODERATE_DEVIATION"
    elif net_difference > 3:
        return "HIGH_DEVIATION"
    else:  # net_difference < 0
        return "LOW_REQUEST"


def get_enhanced_recommendation(row):
    """Generate enhanced recommendation based on difference, stock, and pending orders.
    
    Args:
        row: DataFrame row with difference, stock, and order information
        
    Returns:
        str: Recommendation text
    """
    difference = row['Difference']
    net_difference = row['Net_Difference'] if 'Net_Difference' in row else row['Net Difference']
    available_stock = row['Available Stock']
    pending_orders = row['Pending Orders']
    
    # If requests match forecast exactly
    if difference == 0:
        return "Current forecast appears adequate"
    
    # If requests are lower than forecast
    elif difference < 0:
        return "Consider reducing forecast"
    
    # If deviation is covered by stock
    elif difference <= available_stock:
        return "Use available stock to fulfill requests"
    
    # If deviation is covered by pending orders
    elif difference <= (available_stock + pending_orders):
        return "Use stock and pending orders to fulfill requests"
    
    # If net difference after stock and orders is moderate
    elif 1 <= net_difference <= 3:
        return "Moderate increase needed after using stock & orders"
    
    # If net difference after stock and orders is high
    elif net_difference > 3:
        return "Significant increase needed after using stock & orders"
    
    # Fallback (shouldn't reach here)
    else:
        return "Review needs considering stock and pending orders"


def get_enhanced_status_for_plant(row, plant):
    """Get status for a specific plant's request.
    
    Args:
        row: DataFrame row
        plant (str): Plant name
        
    Returns:
        str: Status string
    """
    plant_requests = row[f'{plant} Requests']
    adjusted_forecast = row['Adjusted Annual Forecast']
    available_stock = row['Available Stock']
    pending_orders = row['Pending Orders']
    
    # Calculate plant-specific difference
    difference = plant_requests - adjusted_forecast
    
    # Calculate plant-specific net difference
    net_difference = difference - available_stock - pending_orders
    
    # Check if deviation is covered by available stock
    if difference > 0 and difference <= available_stock:
        return "COVERED_BY_STOCK"
    
    # Check if deviation is covered by pending orders
    elif difference > 0 and difference <= (available_stock + pending_orders):
        return "COVERED_BY_ORDERS"
    
    # Otherwise use the standard classification
    elif net_difference == 0:
        return "ACCEPTABLE"
    elif 1 <= net_difference <= 3:
        return "MODERATE_DEVIATION"
    elif net_difference > 3:
        return "HIGH_DEVIATION"
    else:  # net_difference < 0
        return "LOW_REQUEST"


def get_enhanced_recommendation_for_plant(row, plant):
    """Generate enhanced recommendation for a specific plant.
    
    Args:
        row: DataFrame row
        plant (str): Plant name
        
    Returns:
        str: Recommendation string
    """
    plant_requests = row[f'{plant} Requests']
    adjusted_forecast = row['Adjusted Annual Forecast']
    available_stock = row['Available Stock']
    pending_orders = row['Pending Orders']
    
    # Calculate plant-specific difference
    difference = plant_requests - adjusted_forecast
    
    # Calculate plant-specific net difference
    net_difference = difference - available_stock - pending_orders
    
    # If requests match forecast exactly
    if difference == 0:
        return "Current forecast appears adequate"
    
    # If requests are lower than forecast
    elif difference < 0:
        return "Consider reducing forecast"
    
    # If deviation is covered by stock
    elif difference <= available_stock:
        return "Use available stock to fulfill requests"
    
    # If deviation is covered by pending orders
    elif difference <= (available_stock + pending_orders):
        return "Use stock and pending orders to fulfill requests"
    
    # If net difference after stock and orders is moderate
    elif 1 <= net_difference <= 3:
        return "Moderate increase needed after using stock & orders"
    
    # If net difference after stock and orders is high
    elif net_difference > 3:
        return "Significant increase needed after using stock & orders"
    
    # Fallback
    else:
        return "Review needs considering stock and pending orders"