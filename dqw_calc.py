from re import S
import gspread
import json
import numpy as np
import datetime
import time
import sys
#ServiceAccountCredentials：Googleの各サービスへアクセスできるservice変数を生成します。
from oauth2client.service_account import ServiceAccountCredentials 
import config

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
print("Enter month day p1 p2 p3 p4 (pi means pi*10000) >>>",end=" ")
month,day,a,b,c,d=(int(x) for x in input().split())
year =2022
today = datetime.datetime(year=year, month=month, day=day, hour=10)
init = datetime.datetime(year=2021, month=11, day=4, hour=0)
row=(today-init).days+2
inputs=[a,b,c,d]
inputs=np.array(inputs)*10000
# print(inputs[0])
inputs=list(inputs)
# print(type(inputs[0]))
# update data first!
for i in range(row,row+1):
    B="B"
    vals=[]
    for j in range(4):
        time.sleep(0.1)
        col=ord(B)+j
        worksheet.update_cell(i,j+2,int(inputs[j]))
# calculate the difference secondly.
for i in range(row,row+1):
    B="B"
    vals=[]
    for j in range(4):
        time.sleep(0.1)
        col=ord(B)+j
        # before place
        before_p=chr(col)+str(i-1)
        # after_p=chr(col)+str(i)
        # before value
        before=(worksheet.acell(before_p).value)
        try:
            int_bef=int(before)
        except:
            print("error: before value doesn't exist!")
            sys.exit()
        dif=inputs[j]-int_bef
        if(dif<6e6 and dif>0):
            vals.append(dif)
    worksheet.update_cell(i,6,int(np.median(vals)))
    print("finished",i,"F",int(np.median(vals)))