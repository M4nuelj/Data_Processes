import datetime
import pandas as pd

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

# Example usage
dates = ['06/28/2023', '06/29/2023']
Disposals = pd.DataFrame()
Disposals['Dates'] = dates
result = list(map(lambda x : get_lot_from_date(x), Disposals.Dates))
Disposals['Lot'] = result
print(Disposals)

