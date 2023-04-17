#coding: utf-8
import openpyxl
import sys

#メッセージ
import tkinter as tk
import tkinter.messagebox as messagebox

# データのアップロード
# 各種ライブラリインポート
import time
import os
from os.path import expanduser
import datetime
import json


today = datetime.date.today().strftime('%Y%m%d')

filename='サンプル_'+today+'.xlsx'
json_str='{\
    "data":[\
    {"column":25,"value":"オプションA"},\
    {"column":26,"value":"オプションB"},\
    {"column":27,"value":"オプションC"},\
    {"column":28,"value":"オプションD"},\
    {"column":29,"value":"オプションE"}\
    ],\
    "filepath":"'+filename+'",\
    "folderpath":""}'

json_dict = json.loads(json_str)


# csv、エクセル　ファイルパス
xlsx_file = '{}'.format(json_dict['filepath'])
print(xlsx_file)

folder = '{}'.format(json_dict['folderpath'])
print(folder)

# 「/」を「￥」に返還
folder_win=folder.replace('/','\\')

#-------------------------------------------------------------

# エクセルファイルを読み込む
wb=openpyxl.load_workbook(xlsx_file)
#sheetを読み込む
ws = wb.worksheets[0]
# エクセルファイルに書き込む
title_str = '{}'.format(json_dict['data'])

# 「'」を「"」に変えないと変換できないので
title_str=title_str.replace('\'','\"')
print(title_str)
title_list = json.loads(title_str)

i=0
for d in title_list:
    print('=======================================')
    print(d)
    ws.cell(row=1,column=d['column'],value=d['value'])

#　保存
wb.save(xlsx_file)
time.sleep(5)

#ファイルを閉じる
wb.close()

quit()
