import xlsxwriter
import xlrd
import pandas as pd
import time
import os

df = pd.read_excel('/Users/mac/Desktop/repos/youtube-comprehensible-input/input_table.xlsx')
print(df.head())

data = {'test': ['hi', 'e'], 'test': ['hi', 'e']}


#Write Excel file
df1 = pd.DataFrame(data, columns= ['test','test'])
df1.to_excel('t.xlsx', index = False, header=True)
