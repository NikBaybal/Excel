import os
import pandas as pd
import openpyxl
import datetime

# Алгоритм обработки данных СШ110кВ
#- привести все строки со временем к формату времени
#- убрать дубликаты
#- привести к UTC
#- создать новый столбец с временем без пропусков
#- заполнить данными исходную таблицу
#- обработать все листы
#- округлить до требуемого значения
#- создать файл csv с требуемым названием

cwd = os.getcwd()
os.chdir("C:/Users/nikip/PycharmProjects/data processing")                                                      # задать путь файла
os.listdir('.')
file = 'СШ110.xlsx'                                                                                             # задать имя файла
diff = datetime.timedelta(hours=9)                                                                              # задать разницу во времени
name='YAGRESNOVAYA'                                                                                             # задать имя станции
xl = pd.ExcelFile(file)
sheets = xl.sheet_names
def data_clean(df_in,num):
    time0 = df_in.columns[0]
    time = df_in.columns[num]
    freq = df_in.columns[num+1]
    df_in = df_in.dropna(subset=[time])
    df_obr = df_in.loc[:, [time, freq]]
    df_obr[time] = pd.to_datetime(df_obr[time],dayfirst=True).round('S')
    df_obr[time] = df_obr[time] - diff
    df_obr[freq] = df_obr[freq].map('{:.3f}'.format)
    df_obr = df_obr.drop_duplicates(subset=time, ignore_index=True)
    df_out = pd.DataFrame(pd.date_range(df_obr[time][0].round('min'), df_obr[time][len(df_obr[time])-1].round('min'), freq="S"), columns=[time])
    df_out = df_out.merge(df_obr, how='left').fillna(method='pad')
    df_out = df_out.fillna(method='bfill')
    begin_data = df_out[time][0].strftime('%Y%m%d')
    begin_time = df_out[time][0].strftime('%H%M%S')
    df_out[time] = df_out[time].dt.strftime('%Y.%m.%d %H:%M:%S')
    df_out = df_out.rename(columns={time: time0})
    return df_out,begin_data,begin_time
def add_csv():
    for row in sheets:
        df_in = xl.parse(row)
        df_out2,begin_data,begin_time=data_clean(df_in,0)
        for row2 in [2,4,6]:
            df_out,begin_data,begin_time= data_clean(df_in,row2)
            df_out2 = df_out2.merge(df_out, how='left')
        df_out2.to_csv(f"{name}.{begin_data}.{begin_time}.csv", index=False, header=False, sep=';')
def main():
    add_csv()

if __name__=='__main__':
    main()




