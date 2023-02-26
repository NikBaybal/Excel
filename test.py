import os
import pandas as pd
import openpyxl
import datetime

# Проверка алгортмов


cwd = os.getcwd()
os.chdir("C:/Users/nikip/PycharmProjects/data processing")
os.listdir('.')
file = 'СШ110.xlsx'
xl = pd.ExcelFile(file)
diff = datetime.timedelta(hours=9)
name='YAGRESNOVAYA'
sheets = xl.sheet_names
df_in = xl.parse('Лист1')
time0= df_in.columns[0]
time = df_in.columns[0]
freq = df_in.columns[1]

df_in = df_in.dropna(subset=[time])
df_obr = df_in.loc[:, [time, freq]]

df_obr[time] = pd.to_datetime(df_obr[time], dayfirst=True).round('S')
df_obr[time] = df_obr[time] - diff
df_obr[freq] = df_obr[freq].map('{:.3f}'.format)
df_obr = df_obr.drop_duplicates(subset=time, ignore_index=True)
df_out = pd.DataFrame(pd.date_range(df_obr[time][0].round('min'), df_obr[time][len(df_obr[time])-1].round('min'), freq="S"), columns=[time])
df_out = df_out.merge(df_obr, how='left').fillna(method='pad')
df_out2 = df_out.fillna(method='bfill')
# begin_data = df_out[time][0].strftime('%Y%m%d')
# begin_time = df_out[time][0].strftime('%H%M%S')
df_out[time] = df_out[time].dt.strftime('%Y.%m.%d %H:%M:%S')
df_out=df_out.rename(columns={time: time0})
print(df_out2)





