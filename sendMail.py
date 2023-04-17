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

today = datetime.date.today().strftime('%Y%m%d')
file_name='OCN光開示用_'+today+'.xlsx'

# PRG2: メール情報の設定
to_address = 'haruna_ochiai@012grp.co.jp,'\
    'rina_takami@012grp.co.jp,'\
    'yuichiro_masuka@012grp.co.jp,'\
    'yu_yamagata@012grp.co.jp,'\
    'shohei_yamaguchi@012grp.co.jp,'\
    'natsumi_soma@012grp.co.jp'
cc_address = 'ryogo_hirano@012grp.co.jp'

# ↓テスト用
#to_address = 'fm_kanri@012grp.co.jp'
#cc_address = 'hiroshi_higashionna@012grp.co.jp'
sendToList = to_address.split(',')+cc_address.split(',')
#
mail_title = "【OCN光達成委員会用_開示データ　" + today
message = '<html>'\
    '<p>お疲れ様です。</P>'\
    '<P></P>'\
    '<P>表題の件でファイルを貼付します。</P>'\
    '</html>'

def logMsg(logMsg):
    #print(logMsg)
    f = open('log.txt', 'a')
    f.write('\n'+datetime.datetime.now().strftime('%Y/%m/%d %H:%M:%S')+'	'+logMsg)
    f.close()
    time.sleep(2)

def sendMailFmKanri(to_address,cc_address,mail_title,message,f_name):
    from_address ='fm_kanri@012grp.co.jp'
    smtp_password = 'LAf2SLEPo1Z5fAx'
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
    try:
        err_msg = '送信中...'
        print(err_msg)
        sendMailFmKanri(to_address,cc_address,mail_title,message,filename)
        time.sleep(5)
    except :
        err_msg = '送信エラー'
        logMsg(err_msg)

if err_msg== '送信中...':
    log_msg = '送信しました。：'+filename
    logMsg(log_msg)
    time.sleep(3)


quit()
