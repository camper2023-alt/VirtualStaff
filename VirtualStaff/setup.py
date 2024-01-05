import sys
from cx_Freeze import setup, Executable
 
base = None
 
if sys.platform == 'win32': base = 'Win32GUI'

# exeにするソースファイルを指定
exe = Executable(script = "virtualstaff.pyw", base= base)
 
setup(name = 'virtualstaff', #ファイルの名前
    version = '1.0',     #バージョン表記
    description = 'KP集計ツール', #アプリケーションの説明
    executables = [exe])  #実行ファイルの形式