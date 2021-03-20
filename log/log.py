from numpy.lib.function_base import _CORE_DIMENSION_LIST
import pandas as pd
import numpy as np
from datetime import datetime

path = 'C:/Users/histi/py37_32/autostock/files/tic_data/'

df = pd.DataFrame()
start_day = '2021-02-26'
date = pd.date_range(start_day, format(datetime.now(), '%Y-%m-%d'))
date = date.astype(str)

# for date in date:
#     print(date)
#     df_new = pd.read_csv(f'{path}{date}.txt', sep=" ",
#                          header=None, encoding='utf-8')

#     df = pd.concat([df, df_new])

df = pd.read_csv(f'{path}test.txt', sep=" ", header=None, encoding='utf-8')
df = pd.read_csv(f'{path}test1.txt', sep=" ", header=None, encoding='utf-8')

df.apply(lambda x: x.str.strip(), axis=1)
df.columns
day = df[0].str.split("[", expand=True)
df['날짜'] = day.drop(0, axis=1)
df.head(1)
col_39 = df[39].str.split("}", expand=True)
df[39] = col_39.drop(1, axis=1)

del_column_list = list(range(12))

df = df.drop(del_column_list, axis=1)
df.head(1)
df = df.drop([12, 14, 16, 18, 20, 22, 24, 26, 28, 30, 32, 34, 36, 38], axis=1)
df.head(1)

df.columns = ['종목명', '스크린번호', 'Fids 번호', '체결시간', '현재가', '전일대비',
              '등락율', '(최우선)매도호가', '(최우선)매수호가', '거래량', '누적거래량', '고가', '시가', '저가', '날짜']

df = df.replace({',': ''}, regex=True)
df.head(1)
df.dtypes
#df1 = df.astype({'현재가': int, '전일대비': int, '등락율': float, '(최우선)매도호가': int, '(최우선)매수호가': int, '거래량': int, '누적거래량': int, '고가': int, '시가': int, '저가': int})
df = df.astype({'현재가': 'int16', '전일대비': 'int16', '등락율': 'float16', '(최우선)매도호가': 'int16',
                '(최우선)매수호가': 'int16', '거래량': 'int16', '누적거래량': 'int32', '고가': 'int16', '시가': 'int16', '저가': 'int16'})
df.dtypes
df['체결시간'] = df['체결시간'].str.replace("'", "")
df['datetime'] = df['날짜']+' ' + df['체결시간']
df['datetime'] = pd.to_datetime(
    df['datetime'], format='%Y-%m-%d %H:%M:%S', errors='raise')
df.dtypes
