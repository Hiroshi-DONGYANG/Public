#coding: utf-8
import openpyxl
import csv
import sys

# データのアップロード
# 各種ライブラリインポート
import selenium
from selenium import webdriver
from selenium.webdriver.support.select import Select
import time
import os
from os.path import expanduser
from webdriver_manager.chrome import ChromeDriverManager
import datetime

if len(sys.argv)<2:
    folder="C:/Users/ユーザー/AppData/Local/Temp/S11"
    recordNo="W＊＊＊＊＊＊＊＊＊＊＊＊"
elif len(sys.argv)==2:
    folder = sys.argv[0]
    recordNo = sys.argv[1]
else :
    pyfilepath=os.path.dirname(__file__)
    os.chdir(pyfilepath)
    folder = sys.argv[1]
    recordNo = sys.argv[2]

import zipfileJPN

print(folder)
print(recordNo)

# csv、エクセル　ファイルパス
csv_file = folder+"/"+recordNo+".csv"
xlsx_file = folder+"/"+r"エクセル.xlsx"


# 「/」を「￥」に返還
folder_win=folder.replace('/','\\')

#このファイルの場所
pyfilepath = __file__

# ダウンロード先を指定
options = webdriver.ChromeOptions()
options.add_experimental_option("prefs", {"download.default_directory": folder_win })

# chromdriver.exe のパスを指定
driver = webdriver.Chrome(ChromeDriverManager().install(), chrome_options=options)

# サイトを開く（ログイン画面）
driver.get("https://a＊＊＊＊＊＊＊＊＊＊")
time.sleep(1)

#ファイルダウンロード
'''
driver.find_element_by_xpath("//div[1]/div[2]/div[2]/form/fieldset/div/div[2]/p[1]/button").click()
time.sleep(3)

# ZIPファイルを読み込み
zip_f = zipfileJPN.ZipFile(folder+'/zipfile.zip')

# ZIPを解凍

zip_f.extractall(folder_win)

# ZIPファイルをクローズ
zip_f.close()
'''

#-------------------------------------------------------------
# エクセルファイルに書き込む
with open(csv_file,newline="") as csvf:
    #CSVファイルを読み込む
    data=csv.reader(csvf)
    # エクセルファイルを読み込む
    wb=openpyxl.load_workbook(xlsx_file)
    #sheetを読み込む
    ws = wb.worksheets[0]
    r=2
    for line in data:
        c=1
        for v in line:
            if c==20:
                v=int(v)
            elif c==40:
                v= datetime.datetime.strptime(v, '%Y/%m/%d')
            
            ws.cell(row=r,column=c).value=v
            c += 1
        r += 1

#　保存
wb.save(xlsx_file)
time.sleep(2)

#ファイルを閉じる
wb.close()

time.sleep(1)

#メッセージ
import tkinter as tk
import tkinter.messagebox as messagebox

tk.Tk().withdraw()
messagebox.showinfo('指示', 'エクセルファイルを開いて「保存」してから閉じて下さい。')


#アップロード
driver.find_element_by_id("upfile").send_keys(xlsx_file )

time.sleep(1)
driver.find_element_by_id("form_submit").click()

#driver.close()
#quit()
