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
#- округлить до требуемого значения
#- создать файл csv с требуемым названием
#- записать все данные в эксель файл

cwd = os.getcwd()
os.chdir("C:/Users/nikip/PycharmProjects/data processing")
os.listdir('.')
file = 'СШ110.xlsx'                                                                               # задать имя файла
xl = pd.ExcelFile(file)
diff = datetime.timedelta(hours=9)                                                               # задать разницу во времени
name='YAGRESNOVAYA'                                                                                             # задать имя станции
sheets = xl.sheet_names
df_in = xl.parse('Лист1')



time = df_in.columns[0]
freq = df_in.columns[1]
df_in=df_in.dropna(subset=[time])
df_in[time] = pd.to_datetime(df_in[time]).round('S')
df_obr = df_in.loc[:,[time, freq]]
df_obr[time] = df_obr[time] - diff
df_obr[freq] = df_obr[freq].map('{:.3f}'.format)
df_obr = df_obr.drop_duplicates(subset=time, ignore_index=True)
df_out = pd.DataFrame(pd.date_range(df_in[time][0] - diff, df_in[time][len(df_in[time])-1] - diff, freq="S"), columns=[time])
df_out = df_out.merge(df_obr, how='left').fillna(method='pad')
begin_data = df_out[time][0].strftime('%Y%m%d')
begin_time = df_out[time][0].strftime('%H%M%S')
df_out[time] = df_out[time].dt.strftime('%Y.%m.%d %H:%M:%S')


print(df_out)










# for row in sheets:
#   df_in = xl.parse(row)
#   for row2 in [0,2,4,6]:
#       df_out, begin_data, begin_time = data_clean(df_in,row2)
#       df_out = df_out.merge(df_out, how='left',on=time)
#   df_out.to_csv(f"{name}.{begin_data}.{begin_time}.csv", index=False, header=False, sep=';')


