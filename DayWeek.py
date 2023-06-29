import datetime

#These 3 functions are able indentify which is the Lot of a product only 
# if the Pack Date is given

def get_week_of_year(date_str):
    date_obj = datetime.datetime.strptime(date_str, '%Y-%m-%d')
    week_number = date_obj.isocalendar()[1]
    return week_number


def get_day_of_week(date_str):
    date_obj = datetime.datetime.strptime(date_str, '%Y-%m-%d')
    day_number = date_obj.weekday() + 2
    return day_number


def get_lot_from_date(date_str): 
    return f'{get_week_of_year(date_str)}{get_day_of_week(date_str)}'


