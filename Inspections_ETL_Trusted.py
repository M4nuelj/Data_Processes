import pandas as pd
import numpy as np
from Inspections_ETL_Raw import input_data2

inspections = []

for sheet_name, sheet_data in input_data2.items():
    for i in range(len(sheet_data['Inspected_by'])):
        inspection = {}
        inspection['Inspected by'] = sheet_data['Inspected_by'][i]
        inspection['Inspection Date'] = sheet_data['Inspection_Date'][i]
        inspection['Lbs Inspected'] = sheet_data['Lbs_Inspected'][i]
        inspection['Lbs Processed'] = sheet_data['Lbs_Processed'][i]
        inspection['Item'] = sheet_data['Item'][i]
        inspection['Product Category'] = sheet_data['Product_Category'][i]
        inspection['WorkOrder'] = sheet_data['WorkOrder'][i]
        inspection['Raw Material'] = sheet_data['Raw_Material'][i]
        inspection['Bags per Box'] = sheet_data['Bags_Box'][i]
        inspection['Label Verification'] = sheet_data['Label_Verification'][i]
        inspection['Tape'] = sheet_data['Tape'][i]
        inspection['Quality Seal'] = sheet_data['Quality_Seal'][i]
        inspections.append(inspection)

Master_Inspections = pd.DataFrame(inspections)

Master_Inspections['Item'] = Master_Inspections['Item'].str.upper()
Master_Inspections['Raw Material'] = Master_Inspections['Raw Material'].str.upper()
Master_Inspections['Bags per Box'] = Master_Inspections['Bags per Box'].str.capitalize()
Master_Inspections['Label Verification'] = Master_Inspections['Label Verification'].str.capitalize()
Master_Inspections['Tape'] = Master_Inspections['Tape'].str.capitalize()
Master_Inspections['Quality Seal'] = Master_Inspections['Quality Seal'].str.capitalize()

Master_Inspections.replace('', np.nan, inplace = True)

Master_Inspections.dropna(inplace = True)

Master_Inspections.to_excel('Master_Inspections.xlsx', index = False, sheet_name = 'Data_Inspections')
