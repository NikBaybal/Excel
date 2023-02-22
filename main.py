import os
import pandas as pd
import openpyxl
import datetime

# План
# -создать новый столбец с временем без пропусков
# -заполнить даныыми исходную таблицу


cwd = os.getcwd()  # Retrieve current working directory (`cwd`)
os.chdir("/content")  # Change directory
os.listdir('.')  # List all files and directories in current directory
file = '20230215 Данные ГТУ1.xlsx'
xl = pd.ExcelFile(file)
df_in = xl.parse('Лист1')


def data_clean(df_in):
    time = df_in.columns[0]
    power = df_in.columns[1]
    freq = df_in.columns[2]
    df_out = df_in[[time, freq, power]]
    df_out[time] = pd.to_datetime(df_out[time]).round('S')
    diff = datetime.timedelta(hours=9)  # разница во времени
    df_out[time] = df_out[time] - diff
    df_out[freq], df_out[power] = df_out[freq].round(3), df_out[power].round(2)
    df_out = df_out.drop_duplicates(subset=time, ignore_index=True)

    return df_out


# n='01'                                     # номер ГТУ
# data_clean(df_in).to_csv(f"YAGRESNOVAYA.{n}.20230201.123500.csv",index=False,header=False)
# sheets = xl.sheet_names
# i=1
# for row in sheets:
#   df_in = xl.parse(row)
#   data_clean(df_in).to_csv(f"YAGRESNOVAYA.{n}.20230201.123500{i}.csv",index=False,header=False)
#   i+=1


df_out = data_clean(df_in)

# print(df_out)
# print(df_out.info())
print(pd.date_range(df_out['time'][0], df_out["time"][20], freq="D"))
# print(df)


pd.date_range(calendar["checkin_date"][0], calendar["checkout_date"][0])
# вывод
DatetimeIndex(['2022-06-01', '2022-06-02', '2022-06-03',
               '2022-06-04', '2022-06-05', '2022-06-06',
               '2022-06-07'],
              dtype='datetime64[ns]', freq='D')
