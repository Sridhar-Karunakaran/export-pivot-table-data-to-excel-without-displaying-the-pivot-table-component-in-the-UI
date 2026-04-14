# Export Pivot Table Data to Excel Without Displaying the Component in the UI

## Repository Description
This repository contains a comprehensive ASP.NET Core application that demonstrates the implementation of exporting Pivot Table data directly to Excel format while maintaining a clean, component-free user interface. The solution showcases how to export complex data structures without rendering the Pivot Table component on the UI, providing a seamless data export experience.

## Project Overview
The Export Pivot Table Data to Excel application is built on ASP.NET Core and provides a robust framework for handling pivot table operations and Excel exports. This project is ideal for developers who need to export structured data to Excel without the overhead of displaying interactive pivot table components on the web interface.

## Features
- **Direct Excel Export**: Export pivot table data directly to Excel files without rendering UI components
- **Data Processing**: Efficient data aggregation and pivot operations using C#
- **Multiple Data Sources**: Support for CSV and JSON data formats
- **Automated Services**: Background services for scheduled data processing tasks
- **RESTful API**: Controller-based endpoints for data management and export operations

## Prerequisites
- .NET 7.0 or later
- Visual Studio 2022 or Visual Studio Code
- Basic knowledge of ASP.NET Core and C#

## Installation
Follow these steps to set up and run the application:

1. **Open the Solution**: Open the `PivotController.sln` file in Visual Studio
2. **Restore Dependencies**: Dependent packages will be downloaded automatically from nuget.org
3. **Build the Application**: Build the solution to ensure all dependencies are resolved
4. **Run the Application**: Press F5 or click the Run button to start the application

## Configuration
The application uses two configuration files:
- `appsettings.json`: Default application settings
- `appsettings.Development.json`: Development-specific configuration overrides

## API Endpoints
The `PivotController` provides endpoints for:
- Retrieving pivot table data
- Exporting data to Excel format
- Managing data sources and transformations

## Data Sources
The application supports multiple data formats:
- **CSV Files**: Located in the `DataSource` folder (e.g., sales.csv)
- **JSON Files**: Structured data files (e.g., sales-analysis.json)

## License
This project is provided as-is for educational and commercial use.
