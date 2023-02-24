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

diff = datetime.timedelta(hours=9)                                              # разница во времени
def data_clean(df_in):
    df_obr = df_in[[time, freq, power]]
    df_obr[time] = pd.to_datetime(df_obr[time]).round('S')
    df_obr[time] = df_obr[time] - diff
    df_obr[freq], df_obr[power] = df_obr[freq].round(3), df_obr[power].round(2)
    df_obr = df_obr.drop_duplicates(subset=time, ignore_index=True)
    return df_obr


# n='01'                                     # номер ГТУ
# data_clean(df_in).to_csv(f"YAGRESNOVAYA.{n}.20230201.123500.csv",index=False,header=False)
# sheets = xl.sheet_names
# i=1
# for row in sheets:
#   df_in = xl.parse(row)
#   data_clean(df_in).to_csv(f"YAGRESNOVAYA.{n}.20230201.123500{i}.csv",index=False,header=False)
#   i+=1

df_obr = data_clean(df_in)
df_out=pd.DataFrame(pd.date_range(df_in[time][0]-diff, df_in[time][len(df_in[time])-1]-diff, freq="S"),columns=[time])
df_out=df_out.merge(df_obr,how='left').fillna(method ='pad')

print(df_out)





