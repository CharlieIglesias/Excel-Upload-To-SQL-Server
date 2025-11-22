# Excel → SQL Scheduled Refresh of Team Task Data

## Overview

This Excel macro automates refreshing departmental tables and writing the data into a secure **`Task_Data` SQL Server** database. It is used daily by **60–80 team members** to maintain a centralised repository of operational data, which is later collated in Excel to calculate productivity, utilisation, and capacity metrics. The solution ensures safe, concurrent usage and reliable data transfer, supporting accurate reporting across the department.

## Key Features & Benefits

- **Automated refresh and write:** Refreshes all relevant tables in Excel, waits for data to stabilise, and writes records directly to the `Task_Data` SQL Server database.  
- **Time-controlled scheduling:** Runs at a predefined time each day, reducing the need for manual intervention.  
- **Safe and concurrent usage:** Prevents conflicts from multiple users running the macro simultaneously, preserving data integrity.  
- **Secure storage:** Centralises departmental data in a SQL Server database for consistency and accessibility.  

## Usage

1. Open the Excel workbook containing the macro.  
2. The macro automatically schedules a daily refresh at the defined time.  
3. When triggered, it:  
   - Refreshes all tables in the workbook.  
   - Waits **10 seconds** for data to stabilise.  
   - Writes the table `TeamDataTable` to the `Task_Data` SQL Server database.  
4. The database is then available to pull into other Excel workbooks for productivity and utilisation calculations.  

## Technologies / Tools

- Excel VBA (Macros)  
- Microsoft SQL Server (`Task_Data`)  

## Impact

- Eliminates manual copy-paste processes for centralising data.  
- Ensures data integrity when writing to SQL, even with multiple daily users.  
- Supports downstream productivity, utilisation, and capacity reporting efficiently.  
- Provides a robust foundation for departmental performance monitoring and decision-making.

