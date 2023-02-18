import os
import pandas as pd
import openpyxl

# Retrieve current working directory (`cwd`)
cwd = os.getcwd()
# Change directory
os.chdir("C:/Users/nikip/PycharmProjects/data processing")
# List all files and directories in current directory
os.listdir('.')
file = 'ГТУ3.xlsx'
# Загружаем spreadsheet в объект pandas
xl = pd.ExcelFile(file)

# Загрузить лист в DataFrame по его имени: df1
df1 = xl.parse('Лист1')

#указывает номер строик со временем
list1=[df1.columns[0],df1.columns[2]]
df1[list1]=df1[list1].astype(str)

for row in list1:
    df1[row]=df1[row].str[1:19]
#print(list(map(lambda x:df1[x].str[1:19],list1)))

df1=df1.drop_duplicates(subset=list1,ignore_index=True)
print(df1)
#print(df1.info())