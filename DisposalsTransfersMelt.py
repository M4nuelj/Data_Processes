import pandas as pd
import datetime

path1 = 'C:/Users/j.renza/Documents/Returns/'

df= pd.read_excel(path1 + 'Disposal & Transfer Log (Responses).xlsx',
                    usecols= ['Type of Return', 'Inspection Date', 'Transfer No.', 'Item',
                     'Quantity (CS)', 'Quantity (Lbs)', 'Pack Date','Lot', 'Vendor',
                     'Reason', 'Comments'])

df2 = pd.read_excel(path1 + 'Products.xlsx')

df3 = df.merge(df2, left_on = 'Item', right_on = 'Item_Code')

def get_week_of_year(date_str):
    date_obj = datetime.datetime.strptime(date_str, '%m/%d/%Y')
    week_number = date_obj.isocalendar()[1]
    return week_number

def get_day_of_week(date_str):
    date_obj = datetime.datetime.strptime(date_str, '%m/%d/%Y')
    day_number = date_obj.weekday() + 2
    return day_number

def get_lot_from_date(date_str):
    return f'{get_week_of_year(date_str)}{get_day_of_week(date_str)}'


Disposals = df3[df3['Type of Return'] == 'Disposal']

Disposals.columns = ['Type', 'Date', 'Transfer', 'Item', 'CS', 'Lbs', 'PackDate', 'Lot', 'Vendor', 'Reason', 'Comments', 'Item_Code', 'Item_Name']

Transfers = df3[df3['Type of Return'] == 'Transfer']

Transfers.columns = ['Type', 'Date', 'Transfer', 'Item', 'CS', 'Lbs', 'PackDate', 'Lot', 'Vendor', 'Reason', 'Comments', 'Item_Code', 'Item_Name']

newDisposals = [Lot if Lot != 531.0 else 'N/A' for Lot in Disposals.Lot]

newTransfers = [Lot if Lot != 531 else 'N/A' for Lot in Transfers.Lot]

newDisposalsPackDate = [PackDate if PackDate != 1/1/2000 else 'N/A' for PackDate in Disposals.PackDate]

newTransfersPackDate = [PackDate if PackDate != 1/1/2000 else 'N/A' for PackDate in Transfers.PackDate]

DisposalsDetails = {'Date': Disposals.Date,
                    'Item': Disposals.Item,
                    'WH01': "X",
                    'WH94': "",
                    'Description': Disposals.Item_Name,
                    'Vendor': Disposals.Vendor,
                    'Disposal': 'D',
                    'Reprocess': "",
                    'Repack': "",
                    'WH02': "",
                    'WH10': "",
                    'WH90': "X",
                    'CS': Disposals.CS,
                    'Weight': Disposals.Lbs,
                    'PackDate': newDisposalsPackDate,
                    'Lot': newDisposals,
                    'Reason': Disposals.Reason,
                    'Comments': Disposals.Comments,
                    'TransferNo': Disposals.Transfer}


TransfersDetails = {'Date': Transfers.Date,
                    'Item': Transfers.Item,
                    'WH01': "X",
                    'WH94': "",
                    'Description': Transfers.Item_Name,
                    'Vendor': Transfers.Vendor,
                    'Transfers': 'T',
                    'Reprocess': "X",
                    'Repack': "",
                    'WH02': "",
                    'WH10': "X",
                    'WH90': "",
                    'CS': Transfers.CS,
                    'Weight': Transfers.Lbs,
                    'PackDate': newTransfersPackDate,
                    'Lot': newTransfers,
                    'Reason': Transfers.Reason,
                    'Comments': Transfers.Comments,
                    'TransferNo': Transfers.Transfer}

FinalDisposals = pd.DataFrame(DisposalsDetails)

FinalTransfers = pd.DataFrame(TransfersDetails)

FinalDisposals.Date = FinalDisposals.Date.dt.strftime('%m/%d/%Y')

FinalDisposals.PackDate = FinalDisposals.PackDate.dt.strftime('%m/%d/%Y')

FinalTransfers.Date = FinalTransfers.Date.dt.strftime('%m/%d/%Y')

FinalTransfers.PackDate = FinalTransfers.PackDate.dt.strftime('%m/%d/%Y')

FinalDisposals['Lot2'] = list(map(lambda x : get_lot_from_date(x), FinalDisposals.PackDate)) 
FinalTransfers['Lot2'] = list(map(lambda x : get_lot_from_date(x), FinalTransfers.PackDate))

print(('File saved sucessfully'))

with pd.ExcelWriter(path1 + 'Disposals & Transfers.xlsx') as writer:
    FinalDisposals.to_excel(writer, sheet_name = 'Disposals', index = False)
    FinalTransfers.to_excel(writer, sheet_name = 'Transfers', index = False)
