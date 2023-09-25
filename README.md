# Python-Data-Wrangling

**Objective**: Validate sales data from Merchandise Management System (MMS) to Data Warehouse (DWH)

Given two separate Excel Files that contains data for MMS and DWH, determine whether there's a variance between the
data source database (MMS) and destination (Data Warehouse). The program must cater all the possible format of data.

DWH Excel file contains: 
1. Sales Sheet 
   - Date in short date format (DD/MM/YYYY)
   - Store Code
   - Gross Sales (no more transformation) or Extended Price (needs grouping transformation)

MMS Excel file contains (can be in separate sheets or in one merged table): 
1. Line Sales Sheet 
   - Transaction date in format of concatenated digits (YYMMDD) or Date in short date format (DD/MM/YYYY)
   - Store Code
   - Line Sales
     
2. Header Sales Sheet 
   - Transaction date in format of concatenated digits (YYMMDD) or Date in short date format (DD/MM/YYYY)
   - Store Code
   - Header Sales 

3. Line Sales Sheet 
   - Transaction date in format of concatenated digits (YYMMDD) or Date in short date format (DD/MM/YYYY)
   - Store Code
   - Line Sales

ACTION TO BE TAKEN: 



