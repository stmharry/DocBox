import datetime
import pandas as pd
import pyodbc

connection_old = pyodbc.connect(
    driver='{Microsoft Access Driver (*.mdb, *.accdb)}',
    dbq=r'C:\Users\警務員\Desktop\1060101獎懲系統-保7版',
)
df_old = pd.read_sql(
    (
        'SELECT * FROM 獎懲登記 '
        'WHERE 識別碼 BETWEEN 20444 AND 21057'
    ),
    con=connection_old,
    index_col='識別碼',
    parse_dates=False,
)

connection_new = pyodbc.connect(
    driver='{Microsoft Access Driver (*.mdb, *.accdb)}',
    dbq=r'C:\Users\Public\DocBox\database.accdb',
)
df_new = pd.read_sql(
    'SELECT TOP 1 * FROM 登記查詢',
    con=connection_new,
)

id_ = df_new['識別碼'].iloc[0] + 1
case_id = df_new['案件編號'].iloc[0] + 1

df_temp = pd.DataFrame()
flag = False
for row in df_old.itertuples(index=False):
    if row['姓名'] is None:
        flag = True
        continue
    elif flag:
        flag = False
        case_id += 1

    id_ += 1

    df_temp.append([{
        '識別碼': id_,
        '案件編號': case_id,
        '姓名': row['姓名'],
        '發文日期': None,
        '發文號': None,
        '結果代碼': None,
        '事由': None,
        '事由代碼': None,
        '法令': None,
        '說明單位': None,
        '說明日期': None,
        '說明文件號': None,
        '說明文件': None,
    }])