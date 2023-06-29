import pandas as pd

# set the date threshold
date_threshold = input('Date: ')

# load the input file into a dictionary of DataFrames, one per sheet
input_file = 'Delta Inspections.xlsx'
input_data = pd.read_excel(input_file, sheet_name = None)

# create a Pandas ExcelWriter to write the output to a new sheet
output_file = 'Delta_Inspections.xlsx'
output_writer = pd.ExcelWriter(output_file, engine = 'openpyxl')

# loop through each sheet in the input file
for sheet_name, sheet_data in input_data.items():
    # select only the rows with dates less than the threshold
    selected_rows = sheet_data[sheet_data['Inspection_Date'] >= date_threshold]

    # write the selected rows to a new sheet in the output file
    selected_rows.to_excel(output_writer, sheet_name = sheet_name, index = False)

# save the output file
output_writer.save()

input_file2 = 'Delta_Inspections.xlsx'
input_data2 = pd.read_excel(input_file2, sheet_name = None)
