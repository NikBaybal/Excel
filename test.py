import os
import pandas as pd
import openpyxl
import datetime

# Алгоритм обработки данных
#- привести все строки со временем к формату времени
#- убрать дубликаты
#- привести к UTC
#- создать новый столбец с временем без пропусков
#- заполнить данными исходную таблицу
#- обработать все листы
#- создать файл csv с требуемым названием

cwd = os.getcwd()                                                               # Retrieve current working directory (`cwd`)
os.chdir("C:/Users/nikip/PycharmProjects/data processing")                      # Change directory
os.listdir('.')                                                                 # List all files and directories in current directory
file = 'ГТУ3.xlsx'
xl = pd.ExcelFile(file)
df_in = xl.parse('Лист1')
time = df_in.columns[0]
power = df_in.columns[1]
freq = df_in.columns[2]
num_min=datetime.timedelta(minutes=5)                                           # количество минут
diff = datetime.timedelta(hours=9)                                              # разница во времени

print(df_in[time][len(df_in[time])-1])
