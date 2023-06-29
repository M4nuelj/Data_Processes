import pandas as pd
import numpy as np
import warnings

warnings.filterwarnings("ignore")

file_path = "C:/Users/j.renza/Documents/Inspections_2023/Delta_Inspections.xlsx"


def read_excel_file(file_path):
    # Open the Excel file
    excel_file = pd.ExcelFile(file_path)

    # Create a dictionary comprehension to store DataFrames for each sheet
    dataframes = {sheet_name: excel_file.parse(sheet_name) for sheet_name in excel_file.sheet_names}

    return dataframes



# Call the function to read the Excel file
result = read_excel_file(file_path)

Templates = ['1A Product Inspection - Equipme',
              '1B Daily inspection sheet- MENU',
              '2A Daily inspection sheet Blend',
              '2B Daily inspection sheet (Blen',
              '3 Daily inspection sheet - Manu',
              '4 Daily inspection sheet - Repa'
              ]

def clean_dictionary(result, Templates):
    cleaned_dict = {}
    for sheet in Templates:
        if sheet in result:
            cleaned_dict[sheet] = result[sheet]
    return cleaned_dict

# Call the fuction to clean the dictionary so it contains only the sheets that match the inspection's templates
result2 = clean_dictionary(result, Templates)

Temp_Columns_1A = ['Title Page_Inspection by', 'Title Page_Inspection date and time', 'Title Page_Product category', 'Title Page_Product item code', 'Title Page_Workorder', 'Title Page_Product Specifications Inspection_Lbs processed', 'Title Page_Product Specifications Inspection_Lbs inspected']
Temp_Columns_1B = ['Title Page_Inspected by', 'Title Page_Inspection date ', 'Title Page_Product category', 'Title Page_Product code', 'Title Page_Workorder', 'Title Page_Product Specifications Inspection_Lbs processed', 'Title Page_Product Specifications Inspection_Lbs inspected']
Temp_Columns_2A = ['Title Page_Inspected by', 'Title Page_Inspection date and time ', 'Title Page_Product category', 'Title Page_Product code', 'Title Page_Workorder', 'Title Page_Product Specifications Inspection_Lbs processed', 'Title Page_Product Specifications Inspection_Lbs inspected']
Temp_Columns_2B = ['Title Page_Inspection by', 'Title Page_Inspection date and time', 'Title Page_Product category', 'Title Page_Product code', 'Title Page_Workorder', 'Title Page_Product Specifications Inspection _Lbs processed', 'Title Page_Product Specifications Inspection _Lbs inspected']
Temp_Columns_3 = ['Title Page_Inspection by', 'Title Page_Inspection date', 'Title Page_Product Category ', 'Title Page_Processed item code', 'Title Page_Workorder', 'Title Page_Product Specifications Inspection _Lbs processed', 'Title Page_Product Specifications Inspection _Lbs inspected']
Temp_Columns_4 = ['Title Page_Inspection by', 'Title Page_Inspection date and time', 'Title Page_Product category', 'Title Page_Product item code', 'Title Page_Workorder', 'Title Page_Product Specifications Inspections_Lbs processed', 'Title Page_Product Specifications Inspections_Lbs inspected']


DFs = []

for key,values in result2.items():
    if key == Templates[0]:
        df_1A = result2[key][Temp_Columns_1A]
        df_1A.columns = ['Inspected_by', 'Date', 'Product_Category', 'Item_Code', 'Workorder', 'Processed_Lbs', 'Inspected_Lbs']
        DFs.append(df_1A)
    elif key == Templates[1]:
        df_1B = result2[key][Temp_Columns_1B]
        df_1B.columns = ['Inspected_by', 'Date', 'Product_Category', 'Item_Code', 'Workorder', 'Processed_Lbs', 'Inspected_Lbs']
        DFs.append(df_1B)
    elif key == Templates[2]:
        df_2A = result2[key][Temp_Columns_2A]
        df_2A.columns = ['Inspected_by', 'Date', 'Product_Category', 'Item_Code', 'Workorder', 'Processed_Lbs', 'Inspected_Lbs']
        DFs.append(df_2A)
    elif key == Templates[3]:
        df_2B = result2[key][Temp_Columns_2B]
        df_2B.columns = ['Inspected_by', 'Date', 'Product_Category', 'Item_Code', 'Workorder', 'Processed_Lbs', 'Inspected_Lbs']
        DFs.append(df_2B)
    elif key == Templates[4]:
        df_3 = result2[key][Temp_Columns_3]
        df_3.columns = ['Inspected_by', 'Date', 'Product_Category', 'Item_Code', 'Workorder', 'Processed_Lbs', 'Inspected_Lbs']
        DFs.append(df_3)
    else:
        df_4 = result2[key][Temp_Columns_4]
        df_4.columns = ['Inspected_by', 'Date', 'Product_Category', 'Item_Code', 'Workorder', 'Processed_Lbs', 'Inspected_Lbs']
        DFs.append(df_4)

for df in DFs:
    df.replace(' ', np.nan, inplace = True)
    df.dropna(inplace = True)
    df.loc[:,'Item_Code'] = df.loc[:,'Item_Code'].str.upper()
    df.loc[:,'Date'] = df.loc[:,'Date'].dt.strftime("%m/%d/%Y")
    df.loc[:,'Workorder'] = df.loc[:,'Workorder'].astype('int')
    df.loc[:,'Workorder'] = df.loc[:,'Workorder'].astype('str')

Total_Inspections = pd.concat(DFs, axis=0)

Total_Inspections.reset_index(drop = True, inplace = True)

try:
    Total_Inspections.to_excel('Total_Inspections.xlsx', index = False)
except PermissionError:
    print("The file 'Total_Inspections.xlsx' is open. Please close it and run the code again.")
else:
    print("File saved successfully.")



    
    
