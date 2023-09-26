# Importing libraries; pandas for data wrangling, glob used for file path, and os for naming the resulting excel file
import pandas as pd
import glob
import os 

# Note for leading apostrophe in Excel
print("\n Note: Remove ' in column headers")

# Excel FILE name format for MMS and DWH <Business_Unit>_<MONTH><YEAR>_<Database> 

# Setting the worksheet, sheet, and column names
MMS_SHEETNAME = 'GROSS_SALES'
COMBINED_DWH = 'DWH_COMBINED'
AGG_SHEETNAME = 'AGG_SALES'
HEADER_SHEETNAME = 'HEADER_SALES'
LINE_SHEETNAME = 'LINE_SALES'
DATE = 'DATE_FLD'
TRANS_DATE = 'TRANS_DATE_(YYMMDD)'
EXTENDED_PRICE = 'EXTENDED_PRICE'
STORE_CODE = 'STORE_CODE'
AGG_COLUMN = 'AGG_SALES'
HEADER_SALES_COLUMN = 'HEADER_SALES'
LINE_SALES_COLUMN = 'LINE_SALES'

# The location of separate xlsx worksheets; Insert additional backslash and add *.xlsx
location = 'C:\\Validation\\Unvalidated\\*.xlsx'

# Create list consisting paths of excel files
unvalidated_files = glob.glob(location)
print('\n Path of Files for Validation: ')
print(unvalidated_files)

# Dataframe for MMS
df_MMS = pd.DataFrame(pd.read_excel(unvalidated_files[1],sheet_name = MMS_SHEETNAME))

# Initiate a variable that indicates whether to summarize or not 
summarize = 0

# Identify whether data needs to be summarized and make a date column if trans date is given
for column in df_MMS.columns:

    if column == EXTENDED_PRICE: 
        summarize = 1 

    elif column == DATE:
        df_MMS[DATE] = pd.to_datetime(df_MMS[DATE],format= '%d/%m/%Y')

    elif column == TRANS_DATE:
        df_MMS[TRANS_DATE] = df_MMS.apply(lambda x: str(x[TRANS_DATE].astype('int')),axis=1)
        df_MMS[DATE] = df_MMS.apply(lambda x: 
                                    (x[TRANS_DATE][-2:] + '/' + x[TRANS_DATE][2:4] + '/' + "20" + x[TRANS_DATE][:2] )
                                    ,axis = 1)
        df_MMS[DATE] = pd.to_datetime(df_MMS[DATE],format= '%d/%m/%Y')

# Summarize the sheet by date and sum of Extended Price when Extended Price is given
if summarize:
    df_MMS = df_MMS.groupby([DATE,STORE_CODE],as_index = False).agg({'EXTENDED_PRICE':'sum'}).round(2)
    df_MMS.columns = ['DATE_FLD', 'STORE_CODE', 'GROSS_SALES']

# IF DWH only have one sheet for 3 tables then proceed with merging
try:
    df_DWH = pd.DataFrame(pd.read_excel(unvalidated_files[0],COMBINED_DWH))
except: ValueError
else: 
    df_DWH[DATE]  = pd.to_datetime(df_DWH[DATE],format= '%d/%m/%Y')
    df_MMS = df_MMS.merge(df_DWH,how = 'left',on=[DATE,STORE_CODE])

# DWH's separate sheets will be red and merged with MMS
DWH_sheets = [AGG_SHEETNAME, HEADER_SHEETNAME,LINE_SHEETNAME]
for sheet in DWH_sheets:
    df_DWH = pd.DataFrame(pd.read_excel(unvalidated_files[0],sheet))
    df_DWH[DATE]  = pd.to_datetime(df_DWH[DATE],format= '%d/%m/%Y')
    df_MMS = df_MMS.merge(df_DWH,how = 'left',on=[DATE,STORE_CODE])

# Calculating Variance 
df_MMS['MMS-AGG'] = df_MMS.apply(lambda x: x['GROSS_SALES'] - x['AGG_SALES'],axis=1).round(2)
df_MMS['MMS-HEADER'] = df_MMS.apply(lambda x: x['GROSS_SALES'] - x['HEADER_SALES'],axis=1).round(2)
df_MMS['MMS-LINE'] = df_MMS.apply(lambda x: x['GROSS_SALES'] - x['LINE_SALES'],axis=1).round(2)

# Main Table of merged MMS and DWH
MMS_VS_DWH = pd.DataFrame(df_MMS[[DATE,STORE_CODE,'GROSS_SALES',
                                 AGG_COLUMN,HEADER_SALES_COLUMN,LINE_SALES_COLUMN,
                                 'MMS-AGG','MMS-HEADER','MMS-LINE']])

# Filtering the Variances
Variance = MMS_VS_DWH[(MMS_VS_DWH['MMS-AGG'] != 0) | (MMS_VS_DWH['MMS-HEADER'] != 0) 
            | (MMS_VS_DWH['MMS-LINE'] != 0)]

print('\n\n*** PREVIEW OF VARIANCE ***\n')
print(Variance)

# Crafting the name of output file
output_file = os.path.basename(unvalidated_files[0][:-8])
output_file = output_file + "MMSvsDWH.xlsx"
output_file = "C:\\Validation\\Validated\\" + output_file

# Double Qoutes needed and name of file at end
writer = pd.ExcelWriter(output_file)

# Writing the sheets for merged MMS and DWH and Variance
MMS_VS_DWH.to_excel(writer,sheet_name='Merged MMS and DWH')
Variance.to_excel(writer,sheet_name='Variance')

# Saving and Closing the File
writer.close()

# Data Validation Complete
print('\n\nData Validation Complete :)\n\n')
