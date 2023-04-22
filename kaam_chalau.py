from data import data_lst, dates_lst, data2_lst
import pandas as pd

df = pd.DataFrame(data2_lst)
print(df.head())

writer = pd.ExcelWriter('pandas_simple.xlsx', engine='xlsxwriter')


df.to_excel(writer, sheet_name='Sheet1')
writer.close()