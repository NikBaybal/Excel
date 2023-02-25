import os
import pandas as pd
import openpyxl
import datetime

# Алгоритм обработки данных
#- привести все строки со временем к формату времени (ошибается если формат времени другой как в данных СШ110кВ)
#- убрать дубликаты
#- привести к UTC
#- создать новый столбец с временем без пропусков
#- заполнить данными исходную таблицу
#- обработать все листы
#- округлить до требуемого значения
#- создать файл csv с требуемым названием
#- записать все данные в эксель файл

cwd = os.getcwd()
os.chdir("C:/Users/nikip/PycharmProjects/data processing")
os.listdir('.')
file = 'ГТУ3.xlsx'                                                                               # задать имя файла
xl = pd.ExcelFile(file)
diff = datetime.timedelta(hours=9)                                                               # задать разницу во времени
def data_clean(df_in):
    time = df_in.columns[0]
    power = df_in.columns[1]
    freq = df_in.columns[2]
    df_in[time] = pd.to_datetime(df_in[time]).round('S')                                    #ошибается если формат времени другой как в данных СШ110кВ)
    df_obr = df_in[[time, freq, power]]
    df_obr[time] = df_obr[time] - diff
    df_obr[freq], df_obr[power] = df_obr[freq].map('{:.3f}'.format), df_obr[power].map('{:.2f}'.format)
    df_obr = df_obr.drop_duplicates(subset=time, ignore_index=True)
    df_out = pd.DataFrame(pd.date_range(df_in[time][0] - diff, df_in[time].iloc[-1] - diff, freq="S"), columns=[time])
    df_out = df_out.merge(df_obr, how='left').fillna(method='pad')
    begin_data = df_out[time][0].strftime('%Y%m%d')
    begin_time = df_out[time][0].strftime('%H%M%S')
    df_out[time] = df_out[time].dt.strftime('%Y.%m.%d %H:%M:%S')
    return df_out,begin_data,begin_time

n='01'                                                                                                          # задать номер ГТУ
name='YAGRESNOVAYA'                                                                                             # задать имя станции
sheets = xl.sheet_names
for row in sheets:
  df_in = xl.parse(row)
  df_out,begin_data,begin_time = data_clean(df_in)
  df_out.to_csv(f"{name}.{n}.{begin_data}.{begin_time}.csv",index=False,header=False,sep=';')








