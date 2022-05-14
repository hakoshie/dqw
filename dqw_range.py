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
print("Enter row number of begin and distance from begin >>>",end=" ")
begin,cnt=(int(x) for x in input().split())

# begin が小さすぎるとき
if(begin<3):
    begin=3

# 前回の行の値
bef_vals=[]

# cnt=0
for i in range(begin,begin+cnt):
    B="B"
    # 差を格納
    difs=[]
    # nowを格納
    nows=[]

    # rowのvaluesを取得
    row_val=worksheet.row_values(i)
    
    time.sleep(0.12)
    for j in range(4):
        # now_p=chr(row)+str(i)
        # now=(worksheet.acell(now_p).value)

        # now=table(i,j) を取得
        try:
            now=int(row_val[j+1])
        except ValueError:
            sys.exit("Currently, there are no data in the row "+str(i)+".") 
        dif=0

        # table(i-1, j)の取得 difの計算
        if(len(bef_vals)==0):
            col=ord(B)+j
            before_p=chr(col)+str(i-1)
            before=(worksheet.acell(before_p).value)
            dif=now-int(before)
        else:
            before=bef_vals[j]
            dif=now-before

        # difが異常でないか
        if(dif<6e6 and dif>0):
            difs.append(dif)
        # nowを次に使うため保存
        nows.append(now)
    # calculate modified median
    worksheet.update_cell(i,6,int(np.median(difs)))
    # message 
    print("Finished",i,"F",int(np.median(difs)))
    # update before values 
    bef_vals=list(nows)

