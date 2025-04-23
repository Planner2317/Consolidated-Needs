# Gasket Inventory Analysis Tool

## Overview

This tool analyzes gasket inventory by comparing master forecast data with plant requests, accounting for available stock and pending orders. It generates a comprehensive Excel report with color-coded status indicators and actionable recommendations.

![Gasket Analysis Dashboard](https://placeholder-for-dashboard-screenshot.png)

## Features

- **Intelligent Deviation Analysis**: Compares plant requests with annual forecasts
- **Inventory Consideration**: Accounts for available stock and pending orders when determining status
- **Enhanced Status Classification**: Uses 6 different status categories with color coding
  - ACCEPTABLE (Green): Perfect match between requests and forecast
  - MODERATE_DEVIATION (Yellow): Difference between 1-3 units
  - HIGH_DEVIATION (Red): Difference > 3 units
  - LOW_REQUEST (Blue): Requests below forecast
  - COVERED_BY_STOCK (Purple): Deviation covered by existing stock
  - COVERED_BY_ORDERS (Pink): Deviation covered by stock + pending orders
- **Smart Recommendations**: Provides specific recommendations based on deviation and inventory status
- **Plant-Specific Analysis**: Individual sheets for plant communication
- **Visual Dashboard**: Summary view highlighting critical deviations and statistics

## Project Structure
