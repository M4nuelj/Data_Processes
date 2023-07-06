import pandas as pd
from DisposalsTransfersMelt import FinalDisposals,FinalTransfers

# This code will create the files with Disposals and Transfers
# It is taking two Data Frames that are created by the DisposalsTransfersMelt code
# Finally one or two files will be exported depending on the existence of Disposals or Transfers

path = 'C:/Users/j.renza/Documents/Returns/'


# The code will ask the user for the transfers number. The user must enter the number or numbers and 
# in case there is more than one they should be separated by ','

DisposalsNumber = input("Type the Transfer Nº for Disposals (if there's more than one separate them by ','): ")

TransfersNumber = input("Type the Transfer Nº for Rework (if there's more than one separate them by ','): ")

DisposalsNumber = DisposalsNumber.split(',')
TransfersNumber = TransfersNumber.split(',')

DNumbers = []
for number in DisposalsNumber:
    DNumbers.append(float(number.strip()))

TNumbers = []
for number in TransfersNumber:
    TNumbers.append(float(number.strip()))

Disposals = FinalDisposals[FinalDisposals.TransferNo.isin(DNumbers)]

Transfers = FinalTransfers[FinalTransfers.TransferNo.isin(TNumbers)]

Disposals = Disposals[['Date', 'TransferNo', 'Item', 'CS', 'Weight', 'PackDate', 'Lot', 'Vendor', 'Reason', 'Comments']]

Transfers = Transfers[['Date', 'TransferNo', 'Item', 'CS', 'Weight', 'PackDate', 'Lot', 'Vendor', 'Reason', 'Comments']]

if Disposals.empty:
    Transfers.to_excel(path + 'Transfers.xlsx', sheet_name = 'Transfers', index = False)
elif Transfers.empty:
    Disposals.to_excel(path + 'Disposals.xlsx', sheet_name = 'Disposals', index = False )
else:
    with pd.ExcelWriter(path + 'Disposal_Transfers.xlsx') as writer:
        Disposals.to_excel(writer, sheet_name = 'Disposals', index = False)
        Transfers.to_excel(writer, sheet_name = 'Transfers', index = False)

print(('File saved sucessfully'))