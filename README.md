# Python-Data-Wrangling

**Objective**: Validate sales data from Merchandise Management System (MMS) anf Data Warehouse (DWH)

Given two separate Excel files that contains data for MMS and DWH in a given time interval, determine whether there's a variance 
between the data in data source database (MMS) and destination (DWH). The program must cater all the possible format of data.

MMS Excel file contains: 
1. Sales Sheet 
   - Transaction date in form of concatenated digits (YYMMDD) or Date in short date format (DD/MM/YYYY)
   - Store Code
   - Gross Sales (no more transformation) or Extended Price (needs grouping transformation)

DWH Excel file contains (can be in separate sheets or in one merged table): 
1. Line Sales Sheet 
   - Date in short date format (DD/MM/YYYY)
   - Store Code
   - Line Sales
     
2. Header Sales Sheet 
   - Date in short date format (DD/MM/YYYY)
   - Store Code
   - Header Sales 

3. Aggregate Sales Sheet 
   - Date in short date format (DD/MM/YYYY)
   - Store Code
   - Aggregate Sales

**ACTION TO BE TAKEN**: 
1. Summarize the Sales table by date and sum of Extended Price in MMS if Extended Price is given, in order to convert it to gross sales.
2. Standardize the format of dates for MMS and DWH if the date format is different.
3. Left join all the tables in DWH (line sales, header sales, and aggregate sales) on MMS using the store code and date.
4. Calculate the variance between gross sales and all the tables in DWH 
5. Export the resulting table and filtered table of variances in an Excel file.
   
**PROGRAM OUTPUT**: An Excel file that contains 2 sheets: 
1. The merged table of MMS and DWH data that shows variance.
2. Filtered stores at a given date that has variance.


