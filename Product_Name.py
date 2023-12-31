import pandas as pd

path = 'C:/Users/j.renza/Documents/Returns/'
df = pd.read_excel(path + 'Products - Copy.xlsx')

df['Name'] = df.Item_Name.str.split(' ')

df['DVendor'] = df.apply(lambda row: 'Montecarlo' if 'MONTECARLO' in row['Name'] else '', axis = 1)

df.to_excel(path + 'Products - Copy.xlsx', index = False)

print('Done')