 モジュールを読み込む
import openpyxl
import msoffcrypto

import datetime
today = datetime.date.today()
today = today.strftime('%Y%m%d')

# パスワード
PASSWORD = "xxxxx"
filename='サンプルファイル'
# 暗号化されたファイル
encrypted_file_name = filename+today+'.xlsx'
# 復号化して保存するファイル
decrypted_file_name = filename+today+'B.xlsx'

# 暗号化ファイルを開く
f = open(encrypted_file_name, "rb")
encrypted_file = msoffcrypto.OfficeFile(f)

# 復号化のパスワードを設定する
encrypted_file.load_key(password=PASSWORD)


# 復号化したファイルを別のファイルに保存する
decrypted_file = open(decrypted_file_name, "wb")
encrypted_file.decrypt(decrypted_file)
f.close()
decrypted_file.close()

# 復号化したファイルを暗号化ファイルの名前に変更
import shutil
shutil.move(decrypted_file_name, encrypted_file_name)

quit()
