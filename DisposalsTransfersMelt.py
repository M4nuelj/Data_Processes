import pandas as pd
import datetime
from DayWeek import get_day_of_week, get_week_of_year, get_lot_from_date

import warnings

warnings.filterwarnings("ignore")

# These are the libraries that will be used in this code

path = 'C:/Users/j.renza/Documents/Returns/'
# path is the location where the files are kept. It is important that the files named
# "Disposal & Transfer Log (Responses)" and "Products"
# are within the same folder as they are both called using the same path

df= pd.read_excel(path + 'Disposal & Transfer Log (Responses).xlsx',
                    usecols= ['Type of Return', 'Inspection Date', 'Transfer No.', 'Item',
                     'Quantity (CS)', 'Quantity (Lbs)', 'Pack Date', 'Vendor',
                     'Reason', 'Comments'])

# df is a DataFrame created from the export of the Google Form found in the following link:
# https://docs.google.com/forms/d/e/1FAIpQLSelAu5GL7vi91SIpMNKC60faZtV_Ei5RJmTBmEJcuqGGQV82Q/viewform

df2 = pd.read_excel(path + 'Products.xlsx')

#df2 is a DataFrame created from an excel file that contains all the items that PrimeMeats has available
# The products are pulled from an external file instead of being named within the same Google Form as a
# way of security... Items can change name;therefore, it'll be easier to re-write the entire file if the names are
# kept in a different souce.

df3 = df.merge(df2, left_on = 'Item', right_on = 'Item_Code')

# df3 is a DataFrame that contains the combination of df and df2. The merge
# is done with the item as a key; therefore, if written information 
# cannot be found in the final file the Products.xlsx file should be inspected.


# All the code below cleans the data and organize it in a way that match the SQF format.

Disposals = df3[df3['Type of Return'] == 'Disposal']

Disposals.columns = ['Type', 'Date', 'Transfer', 'Item', 'CS', 'Lbs', 'PackDate', 'Vendor', 'Reason', 'Comments', 'Item_Code', 'Item_Name', 'Item_Type_Code', 'Item_Type', 'Vendor2']

Transfers = df3[df3['Type of Return'] == 'Transfer']

Transfers.columns = ['Type', 'Date', 'Transfer', 'Item', 'CS', 'Lbs', 'PackDate', 'Vendor', 'Reason', 'Comments', 'Item_Code', 'Item_Name', 'Item_Type_Code', 'Item_Type', 'Vendor2']

Disposals['PackDate'] = Disposals.PackDate.astype('str')
Transfers['PackDate'] = Transfers.PackDate.astype('str')

Disposals['Lot'] = Disposals.apply(lambda row: int(get_lot_from_date(row['PackDate'])) if row['Item_Type'] != 'BIBO' else 'N/A', axis = 1)

Transfers['Lot'] = Transfers.apply(lambda row: int(get_lot_from_date(row['PackDate'])) if row['Item_Type'] != 'BIBO' else 'N/A', axis = 1)

Disposals['NewPackDate'] = Disposals.apply(lambda row: row['PackDate'] if row['PackDate'] != '2000-01-01' else 'N/A', axis = 1)

Transfers['NewPackDate'] = Transfers.apply(lambda row: row['PackDate'] if row['PackDate'] != '2000-01-01' else 'N/A', axis = 1)

DisposalsDetails = {'Date': Disposals.Date,
                    'Item': Disposals.Item,
                    'WH01': "X",
                    'WH94': "",
                    'Description': Disposals.Item_Name,
                    'Vendor': Disposals.Vendor2,
                    'Disposal': 'D',
                    'Reprocess': "",
                    'Repack': "",
                    'WH02': "",
                    'WH10': "",
                    'WH90': "X",
                    'CS': Disposals.CS,
                    'Weight': Disposals.Lbs,
                    'PackDate': Disposals.NewPackDate,
                    'Lot': Disposals.Lot,
                    'Reason': Disposals.Reason,
                    'Comments': Disposals.Comments,
                    'TransferNo': Disposals.Transfer}


TransfersDetails = {'Date': Transfers.Date,
                    'Item': Transfers.Item,
                    'WH01': "X",
                    'WH94': "",
                    'Description': Transfers.Item_Name,
                    'Vendor': Transfers.Vendor2,
                    'Transfers': 'T',
                    'Reprocess': "X",
                    'Repack': "",
                    'WH02': "",
                    'WH10': "X",
                    'WH90': "",
                    'CS': Transfers.CS,
                    'Weight': Transfers.Lbs,
                    'PackDate': Transfers.NewPackDate,
                    'Lot': Transfers.Lot,
                    'Reason': Transfers.Reason,
                    'Comments': Transfers.Comments,
                    'TransferNo': Transfers.Transfer}

FinalDisposals = pd.DataFrame(DisposalsDetails)

FinalTransfers = pd.DataFrame(TransfersDetails)

FinalDisposals.Date = FinalDisposals.Date.dt.strftime('%m/%d/%Y')

FinalTransfers.Date = FinalTransfers.Date.dt.strftime('%m/%d/%Y')


FinalDisposals['PackDate'] = pd.to_datetime(FinalDisposals['PackDate'], errors='coerce')
FinalDisposals['PackDate'] = FinalDisposals['PackDate'].dt.strftime('%m/%d/%Y')
FinalDisposals.loc[pd.isna(FinalDisposals['PackDate']), 'PackDate'] = 'N/A'


FinalTransfers['PackDate'] = pd.to_datetime(FinalTransfers['PackDate'], errors='coerce')
FinalTransfers['PackDate'] = FinalTransfers['PackDate'].dt.strftime('%m/%d/%Y')
FinalTransfers.loc[pd.isna(FinalTransfers['PackDate']), 'PackDate'] = 'N/A'

# The code below saves both Dataframes with the disposals and Transfers Information into a new excel file
# that is called 'Disposals & Transfers'


with pd.ExcelWriter(path + 'Disposals & Transfers.xlsx') as writer:
    FinalDisposals.to_excel(writer, sheet_name = 'Disposals', index = False)
    FinalTransfers.to_excel(writer, sheet_name = 'Transfers', index = False)

print(('File saved sucessfully'))
