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
from sqlalchemy import not_ 

#2つのAPIを記述しないとリフレッシュトークンを3600秒毎に発行し続けなければならない
scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']

#認証情報設定
#ダウンロードしたjsonファイル名をクレデンシャル変数に設定（秘密鍵、Pythonファイルから読み込みしやすい位置に置く）
credentials = ServiceAccountCredentials.from_json_keyfile_name('./pythonsheet.json', scope)

#OAuth2の資格情報を使用してGoogle APIにログインします。
gc = gspread.authorize(credentials)

#共有設定したスプレッドシートキーを変数[SPREADSHEET_KEY]に格納する。
SPREADSHEET_KEY = config.SPREADSHEET_KEY


#共有設定したスプレッドシートのシート1を開く
worksheet = gc.open_by_key(SPREADSHEET_KEY).sheet1

#A1セルの値を受け取る
import_value = int(worksheet.acell('B2').value)

#A1セルの値に100加算した値をB1セルに表示させる
print("Enter start column >>>",end=" ")
col=int(input())

for i in range(col,col+1):
    B="B"
    vals=[]

    for j in range(4):
        time.sleep(0.1)
        column=ord(B)+j
        # before place
        
        # after_p=chr(row)+str(i)
        # before value
        row=i
        not_found=[]
        init_val=int(worksheet.acell(chr(column)+str(row)).value)
        while(row-i<10):
            row+=1
            next_p=chr(column)+str(row)
            next=(worksheet.acell(next_p).value)
            try:
                int_next=int(next)
                break
            except:
                not_found.append(row)
                continue
        dist=row-i
        for enum,row_c in enumerate(not_found):
            worksheet.update_cell(row_c,j+2,int((int_next-init_val)*(enum+1)/(1+len(not_found))+init_val))
            print(chr(column),row_c,int((int_next-init_val)*(enum+1)/(1+len(not_found))+init_val))
begin=col
cnt=row-col
data=[]
label=["row","date","hoge","meu","fuwafuwa","marika","dif(med)","sum", "MA(all)", "MA(7-day)"]
for i in range(begin,begin+cnt):
    # rowのvaluesを取得
    row_val=worksheet.row_values(i)
    row_val.insert(0,i)
    # print(row_val)
    data.append(row_val)
    # time.sleep(0.05)
df=pd.DataFrame(data,columns=label)
print(df.to_string(index=False))
