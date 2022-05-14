import gspread
import json
import numpy as np
import datetime
import time
import sys
import pandas as pd
import config
#ServiceAccountCredentials：Googleの各サービスへアクセスできるservice変数を生成します。
from oauth2client.service_account import ServiceAccountCredentials 

#2つのAPIを記述しないとリフレッシュトークンを3600秒毎に発行し続けなければならない
scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']

#認証情報設定
#ダウンロードしたjsonファイル名をクレデンシャル変数に設定（秘密鍵、Pythonファイルから読み込みしやすい位置に置く）
credentials = ServiceAccountCredentials.from_json_keyfile_name('C:/Users/hakos/Dropbox/Downloads/misc/dqw/pythonsheet.json', scope)

#OAuth2の資格情報を使用してGoogle APIにログインします。
gc = gspread.authorize(credentials)

#共有設定したスプレッドシートキーを変数[SPREADSHEET_KEY]に格納する。
SPREADSHEET_KEY = config.SPREADSHEET_KEY


#共有設定したスプレッドシートのシート1を開く
worksheet = gc.open_by_key(SPREADSHEET_KEY).sheet1

#A1セルの値を受け取る
import_value = int(worksheet.acell('B2').value)

#A1セルの値に100加算した値をB1セルに表示させる
print("Enter column number of begin and distance from begin >>>",end=" ")
begin,cnt=(int(x) for x in input().split())

# begin が小さすぎるとき
if(begin<3):
    begin=3

label=["row","date","hoge","meu","fuwafuwa","marika","dif(med)","sum", "MA(all)", "MA(7-day)"]
# print(label)
data=[]
for i in range(begin,begin+cnt):
    # rowのvaluesを取得
    row_val=worksheet.row_values(i)
    row_val.insert(0,i)
    # print(row_val)
    data.append(row_val)
    # time.sleep(0.05)
df=pd.DataFrame(data,columns=label)
print(df.to_string(index=False))
