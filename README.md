# 📊 Export Pivot Table Data to Excel Without UI Component Display

[![.NET](https://img.shields.io/badge/.NET-7.0-512BD4.svg)](https://dotnet.microsoft.com/)
[![Visual Studio](https://img.shields.io/badge/Visual%20Studio-2022-5C2D91.svg)](https://visualstudio.microsoft.com/)
[![Syncfusion](https://img.shields.io/badge/Syncfusion-Pivot%20Engine-FF6B35.svg)](https://www.syncfusion.com/)
[![Status](https://img.shields.io/badge/status-stable-brightgreen.svg)](#)

> **Server-side pivot table processing with direct Excel export** — Generate professional pivot table exports to Excel files without rendering UI components using ASP.NET Core and Syncfusion Pivot Engine.

## 🔍 Overview

This project demonstrates a production-ready implementation for exporting Pivot Table data directly to Excel format while maintaining a clean, component-free architecture. Perfect for business intelligence dashboards, financial reporting tools, and automated data export pipelines that don't require interactive UI pivot controls.

## ✨ Key Features

- ✅ **Headless Pivot Export**: Export pivot data to Excel without rendering pivot table UI components
- ✅ **Server-Side Processing**: Efficient pivot aggregation using Syncfusion Pivot Engine
- ✅ **Multiple Data Sources**: Support for in-memory collections, CSV, and JSON data formats
- ✅ **Email Integration**: Automated email delivery of Excel exports with attachments
- ✅ **Memory Caching**: High-performance data caching for repeated pivot operations
- ✅ **RESTful API**: Clean REST endpoints for pivot data retrieval and export
- ✅ **Background Services**: Scheduled tasks with timed hosted services
- ✅ **CORS Enabled**: Cross-origin resource sharing for flexible frontend integration

## 🛠 Technology Stack

### Required Technologies
- **Framework**: ASP.NET Core 7.0+
- **IDE**: Visual Studio 2022 or Visual Studio Code
- **Language**: C# 11+
- **Excel Engine**: Syncfusion.XlsIO.Net.Core (v27.1.52+)
- **Pivot Engine**: Syncfusion.Pivot.Engine (v27.1.52+)
- **Data Format**: JSON, CSV, In-Memory Collections

### System Requirements
- **.NET SDK**: 7.0 or higher
- **Runtime**: ASP.NET Core Runtime 7.0+
- **Memory**: 2GB minimum
- **Platform**: Windows, macOS, or Linux

## 📦 Installation & Setup

### Step 1: Clone the Repository
```bash
git clone https://github.com/SyncfusionExamples/export-pivot-table-data-to-excel-without-displaying-the-pivot-table-component-in-the-UI
```

### Step 2: Open in Visual Studio
Open the `PivotController.sln` file in Visual Studio 2022:
```
File → Open → PivotController.sln
```

### Step 3: Restore NuGet Packages
Right-click the Solution and select **Restore NuGet Packages**. Alternatively, use the Package Manager Console:
```bash
Update-Package
```

### Step 4: Build the Solution
Build the entire solution using:
```bash
Build → Build Solution (Ctrl+Shift+B)
```

### Step 5: Configure Email (Optional)
Edit the `SendEMail` method in `PivotController.cs` to add your SMTP credentials:
- Update `client.Host` with your email provider's SMTP server
- Replace `from` and `recipients` email addresses
- Set `client.Credentials` with valid app password

### Step 6: Run the Application
Start the application by pressing **F5** or clicking the Run button.

The application will launch on `http://localhost:5000` or `http://localhost:5001` (HTTPS).

## 🚀 Quick Start

1. **Build and run** the solution (F5)
2. **API endpoint** is available at: `http://localhost:5000/api/pivot/post`
3. **Trigger export** by making a POST request to the endpoint
4. **Excel file** is saved to `D:\Export\Sample.xlsx`
5. **Optional**: Email notification sent with the generated Excel attachment

### Sample POST Request
```csharp
POST /api/pivot/post HTTP/1.1
Host: localhost:5000
Content-Type: application/json

{
  "Action": "onExcelExport",
  "Hash": "a8016852-2c03-4f01-b7a8-cdbcfd820df1",
  "ExportAllPages": true
}
```

## 🗂 Project Structure

```
PivotController/
├── PivotController.csproj          # Project configuration with dependencies
├── PivotController.sln             # Visual Studio solution file
├── Program.cs                      # ASP.NET Core startup configuration
├── appsettings.json                # Application settings
├── appsettings.Development.json    # Development configuration
│
├── Controllers/
│   └── PivotController.cs          # Main API controller for pivot operations
│
├── DataSource/
│   ├── DataSource.cs               # Data source definitions and providers
│   ├── sales.csv                   # Sample CSV data file
│   └── sales-analysis.json         # Sample JSON data file
│
├── Services/
│   └── TimedHostedService.cs       # Background service for scheduled tasks
│
├── Properties/
│   └── launchSettings.json         # Debug launch configuration
│
├── bin/                            # Compiled binaries
└── obj/                            # Build artifacts
```

## 💡 Core Architecture

### Data Flow Diagram
```
User Request (POST /api/pivot/post)
    ↓
PivotController.Post()
    ↓
Load Data Source (CSV/JSON/In-Memory)
    ↓
Configure Pivot Engine Settings
    ↓
Generate Pivot Report
    ↓
Create Excel Workbook via Syncfusion.XlsIO
    ↓
Update Worksheet with Pivot Data
    ↓
Save to File System & Send Email
    ↓
Return Response
```

### Key Components

#### PivotController
- Handles API requests for pivot data and Excel export
- Manages pivot engine configuration and data aggregation
- Implements memory caching for performance optimization
- Triggers email notifications with Excel attachments

#### DataSource.cs
- Provides virtual product data with sales metrics
- Supports multiple data source formats (CSV, JSON, Collections)
- Implements INotifyPropertyChanged for real-time updates

#### TimedHostedService
- Background service for scheduled pivot processing
- Executes periodic data aggregation tasks
- Integrates with ASP.NET Core hosted service framework

## ⚙️ Configuration

### Pivot Engine Settings
Configure pivot field layout in `PivotController.cs`:
```csharp
var param = new FetchData
{
    Rows = new List<FieldOptions>
    {
        new FieldOptions { Name = "ProductID" }
    },
    Columns = new List<FieldOptions>
    {
        new FieldOptions { Name = "Country" }
    },
    Values = new List<FieldOptions>
    {
        new FieldOptions { Name = "Price", Caption = "Price" },
        new FieldOptions { Name = "Sold", Caption = "Units Sold" }
    }
};
```

### Application Settings
Configure in `appsettings.json`:
```json
{
  "Logging": {
    "LogLevel": {
      "Default": "Information",
      "Microsoft.AspNetCore": "Warning"
    }
  },
  "AllowedHosts": "*"
}
```

### Data Source Options
Uncomment the desired data source in `GetData` method:
```csharp
// Option 1: In-memory virtual data (default)
return new DataSource.PivotViewData().GetVirtualData();

// Option 2: CSV file
return new DataSource.PivotCSVData().ReadCSVData("DataSource/sales.csv");

// Option 3: JSON file
return new DataSource.PivotJSONData().ReadJSONData("DataSource/sales-analysis.json");

// Option 4: Remote CSV via HTTP
return new DataSource.PivotCSVData().ReadCSVData("http://cdn.syncfusion.com/data/sales-analysis.csv");
```

## 🎯 Usage Examples

### Export Pivot to Excel
The application automatically exports pivot data on startup. To manually trigger:
```csharp
var controller = new PivotController(cache, environment);
await controller.Post();
```

### Customize Pivot Configuration
Modify the `param` object in `Post()` method to change:
- **Row Fields**: Product categories, regions, time periods
- **Column Fields**: Countries, years, quarters
- **Value Fields**: Sum, Count, Average aggregations

### Enable Server-Side Aggregation
```csharp
EnableServerSideAggregation = true
```

## ❓ Troubleshooting & FAQ

**Q: Application won't start**
- Ensure .NET 7.0 SDK is installed: `dotnet --version`
- Check if port 5000/5001 is available
- Review console output for specific error messages

**Q: Excel file not being created**
- Verify export path exists: `D:\Export\`
- Check file permissions on the export directory
- Ensure Syncfusion.XlsIO NuGet package is properly installed

**Q: Email not sending**
- Configure SMTP credentials in `SendEMail()` method
- Use app-specific password for Gmail/Office365
- Check firewall/proxy settings for port 587 access

**Q: Performance issues with large datasets**
- Enable `EnableServerSideAggregation = true`
- Increase memory cache size in `Program.cs`
- Consider implementing pagination

**Q: Port already in use**
- Change port in `Properties/launchSettings.json`
- Or use PowerShell: `netstat -ano | findstr :5000`

## 📚 Resources

- [ASP.NET Core Documentation](https://learn.microsoft.com/en-us/aspnet/core/)
- [Syncfusion Pivot Engine Guide](https://ej2.syncfusion.com/aspnetcore/documentation/pivot-table/server-side-pivot-engine)
- [Syncfusion Excel Export Guide](https://ej2.syncfusion.com/aspnetcore/documentation/pivot-table/excel-export)
- [MDN Web Documentation](https://developer.mozilla.org/)

## 🤝 Contributing

Contributions are welcome! To contribute:
1. Fork the repository
2. Create a feature branch (`git checkout -b feature/YourFeature`)
3. Commit changes (`git commit -m 'Add YourFeature'`)
4. Push to the branch (`git push origin feature/YourFeature`)
5. Open a Pull Request

## 📄 License

This project is licensed under the **Syncfusion Community License**. See [Syncfusion License](https://www.syncfusion.com/content/downloads/syncfusion_license.pdf) for details.

## 🆘 Support & Issues

For questions, issues, or suggestions:
- 📧 Open a GitHub issue with detailed reproduction steps
- 💬 Review the Troubleshooting section above
- 🔍 Check existing documentation and resource links

---
