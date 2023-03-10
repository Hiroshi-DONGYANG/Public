import pyminizip
import glob
import shutil
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import time
import datetime
today = datetime.date.today().strftime('%Y%m%d')
print(today)

zippassword = 'passwordx'
#
# PRG2: メール情報の設定
to_address = "sample@sample.jp"
cc_address = "sample@sample.co.jp"
mail_title = "【メール件名】"+today
file_name = 'エクセルファイル.xlsx'
message = '''
<html>
    <p>いつもお世話になっております。</p>
    <p>○○でございます。</p>
    <p></p>
    <p>お手数おかけしますが、ご確認の程、お願い申し上げます。</p>
    <p></p>
    <p>zipパスワード：~password~</p>
    <p>■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■</p>
    <p></p>
    <p>株式会社○○</p>
    <p></p>
    <p>このメッセージは自動送信システムを利用しております。</p>
    <p>ご返信は担当窓口までお願い致します。</p>
    <p></p>
    <p>■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■</p>
</html>
'''
message.replace('~password~', zippassword)


def logMsg(logMsg):
    #print(logMsg)
    f = open('log.txt', 'a')
    f.write('\n'+datetime.datetime.now().strftime('%Y/%m/%d %H:%M:%S')+'	'+logMsg)
    f.close()
    time.sleep(2)

def sendMailFmKanri(to_address,cc_address,mail_title,message,f_name):
    from_address ='sousin@sample.co.jp'
    smtp_password = 'smtppassword'
    #
    # メール送信
    try:
        msg = MIMEMultipart()
        msg['Subject'] = mail_title
        msg['To'] = to_address
        msg['Cc'] = cc_address
        msg['From'] = from_address
        msg.attach(MIMEText(message,"html"))
        with open(f_name, "rb") as f:
            attachment = MIMEApplication(f.read())
    except:
        err_msg = 'メール情報作成失敗'
        logMsg(err_msg)
    #
    sendToList = to_address.split(',')+cc_address.split(',')
    attachment.add_header("Content-Disposition", "attachment", filename=f_name)
    msg.attach(attachment)
    #
    obj_smtp= smtplib.SMTP('smtp.gmail.com', 587)
    obj_smtp.starttls()
    obj_smtp.login(from_address, smtp_password)
    #
    obj_smtp.sendmail(from_address, sendToList, msg.as_string())
    obj_smtp.quit()

if len(glob.glob(file_name))==0:
    err_msg = 'CSVファイルがありません。'
    logMsg(err_msg)

for filename in glob.glob(file_name):
    filename_zip = filename.replace('.xlsx', '.zip')
    try :
        pyminizip.compress(filename, "", filename_zip, zippassword, 0)
        try:
            print("送信中...")
            sendMailFmKanri(to_address,cc_address,mail_title,message,filename_zip)
            pritn(to_address)
            pritn(to_address)
            pritn(to_address)
            pritn(to_address)
            time.sleep(5)
        except :
            err_msg = '送信エラー'
            logMsg(err_msg)
    except :
        if err_msg != '送信エラー':
            err_msg = '圧縮エラー'
            logMsg(err_msg)


log_msg = '送信しました。：'+filename_zip
logMsg(log_msg)

quit()
