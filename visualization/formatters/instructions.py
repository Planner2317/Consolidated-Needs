from openpyxl.styles import Font
from config import CURRENT_DATETIME, CURRENT_USER

def create_instructions_sheet(wb):
    """Create and format the instructions sheet.
    
    Args:
        wb: Excel workbook object
    """
    
    instructions = wb.create_sheet('Instructions')
    instructions['A1'] = "GASKET INVENTORY ANALYSIS - INSTRUCTIONS"
    instructions['A1'].font = Font(size=14, bold=True)
    
    # Add sections
    instructions['A3'] = "PURPOSE:"
    instructions['A4'] = "This analysis compares the master forecast data with current plant requests for gasket items, while considering available inventory and pending orders."
    instructions['A4'].alignment = None  # Set wrap_text=True when creating the sheet
    
    instructions['A6'] = "SHEETS:"
    instructions['A7'] = "1. Master: Original forecast data based on 3-year consumption history"
    instructions['A8'] = "2. Responses: Current plant requests data"
    instructions['A9'] = "3. Comparison Analysis: Detailed analysis comparing forecasts with current requests"
    instructions['A10'] = "4. Summary Statistics: Overview metrics of the analysis"
    instructions['A11'] = "5. Plant Summary: Breakdown of requests by plant"
    instructions['A12'] = "6. Dashboard: Visual representation of key insights"
    instructions['A13'] = "7. Plant_[Name]: Individual plant sheets for communication with each plant"
    instructions['A14'] = "8. Instructions: This guide"
    
    instructions['A16'] = "INVENTORY CONSIDERATION:"
    instructions['A17'] = "The analysis now considers:"
    instructions['A18'] = "- Stock Qty: Current available stock"
    instructions['A19'] = "- Open PRs Total 24 Months: Quantities in open purchase requisitions"
    instructions['A20'] = "- Open POs Total 24 Months: Quantities in open purchase orders"
    instructions['A21'] = "- Pr Not Confirmed 24 Months: Quantities in purchase requisitions not yet confirmed"
    instructions['A22'] = "These quantities are used to determine if deviations can be covered by existing inventory or pending orders."
    
    instructions['A24'] = "ENHANCED DEVIATION CLASSIFICATIONS:"
    instructions['A25'] = "- ACCEPTABLE (Green): Total Plant Requests exactly match Annual Forecast (difference = 0)"
    instructions['A26'] = "- MODERATE DEVIATION (Yellow): Total Plant Requests are 1 to 3 units higher than Annual Forecast (difference = 1 to 3)"
    instructions['A27'] = "- HIGH DEVIATION (Red): Total Plant Requests are more than 3 units higher than Annual Forecast (difference > 3)"
    instructions['A28'] = "- LOW REQUEST (Blue): Total Plant Requests are lower than Annual Forecast (difference < 0)"
    instructions['A29'] = "- COVERED BY STOCK (Purple): Deviation can be covered by existing stock"
    instructions['A30'] = "- COVERED BY ORDERS (Pink): Deviation can be covered by stock and pending orders combined"
    
    # Add note about forecast adjustment
    instructions['A32'] = "IMPORTANT NOTE: Annual Forecast values less than 1 are adjusted to 0 for analysis purposes."
    instructions['A32'].font = Font(bold=True, italic=True)
    
    instructions['A34'] = "ENHANCED RECOMMENDATIONS:"
    instructions['A35'] = "- Use available stock to fulfill requests: When deviation can be covered by current stock"
    instructions['A36'] = "- Use stock and pending orders to fulfill requests: When deviation can be covered by stock and pending orders"
    instructions['A37'] = "- Significant increase needed after using stock & orders: When net difference after stock and orders is > 3"
    instructions['A38'] = "- Moderate increase needed after using stock & orders: When net difference after stock and orders is 1-3"
    instructions['A39'] = "- Consider reducing forecast: When Total Plant Requests are lower than Annual Forecast"
    instructions['A40'] = "- Current forecast appears adequate: When difference is exactly 0 (perfect match)"
    
    instructions['A42'] = "PLANT COMMUNICATION SHEETS:"
    instructions['A43'] = "Each plant has its own dedicated sheet (Plant_[Name]) containing:"
    instructions['A44'] = "1. All response data for that plant"
    instructions['A45'] = "2. Key master data fields for requested items"
    instructions['A46'] = "3. Plant-specific deviation analysis including inventory consideration"
    instructions['A47'] = "4. Communication with the Modified User from each plant should reference this sheet"
    
    instructions['A49'] = "HOW TO USE THIS ANALYSIS:"
    instructions['A50'] = "1. Review the Dashboard for a quick overview of key issues"
    instructions['A51'] = "2. Check the Comparison Analysis sheet for detailed item-by-item review"
    instructions['A52'] = "3. Focus on items with HIGH DEVIATION status that cannot be covered by stock or orders"
    instructions['A53'] = "4. For items with deviations, check if they can be fulfilled using available stock or pending orders"
    instructions['A54'] = "5. Use the Plant Summary to understand which plants have the most deviations"
    instructions['A55'] = "6. For plant-specific communication, use the dedicated Plant_[Name] sheets"
    
    instructions['A57'] = f"Generated on: {CURRENT_DATETIME}"
    instructions['A58'] = f"Generated by: {CURRENT_USER}"
    
    # Format sections
    for row in [3, 6, 16, 24, 34, 42, 49]:
        instructions.cell(row=row, column=1).font = Font(bold=True)
    
    # Set column width
    instructions.column_dimensions['A'].width = 100