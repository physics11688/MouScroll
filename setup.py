import subprocess
import re
from platform import system
import sys
import os
from shutil import copy


def color_print(color: int, text: str, is_bold=False, end="\n") -> None:
    """色付きのprint関数

    Args:
        color: 色番号

        text: prints帯文字列

        is_bold: ボールドにしたきゃ True

        end: 改行したくないなら end=""
    """
    if is_bold:
        text = f"\033[1m{text:10s}\033[0m"
    if end == "":
        print(f"\033[{color}m{text}\033[0m", end="")
    else:
        print(f"\033[{color}m{text}\033[0m")


def YorN() -> bool:
    user_input = input("Yes or No >> ")

    Yes = re.search(r"/yes/i", user_input)
    Y = re.search(r"[Yy]", user_input)
    if Yes is Y is None:
        return False
    else:
        return True


msg = """これはMouScroll.pyのインストールスクリプトです。
以下の処理を行います。

1. pipを使って pynput, pywin32, pyautogui, pygetwindowのインストール

2. ホームディレクトリ直下に local\\bin\\MouScroll フォルダを作り、スクリプト MouScroll.pyw を突っ込む



"""

print(msg)

# OSの確認
if system() != 'Windows':
    print("MouScrollは Windowsのみ使用可能です")
    print("プログラムを終了します")
    exit(2)

color_print(35, "セットアップを実行しますか？\n")
if not YorN():
    print("プログラムを終了します")
    exit(1)

try:
    subprocess.run(["pip", "install", "-U", "pynput", "pywin32", "pyautogui", "pygetwindow"])
except:
    print("モジュールのインストールに失敗しました")
    print("プログラムを終了します")

# ディレクトリの移動
os.chdir(os.path.dirname(__file__))
HOME = os.path.expanduser("~")
if " " in HOME:
    print("\nホームディレクトリのPATHに 空白 が含まれています.")
    print("出直してきやがれ.")
    exit(3)
INSTALL_PATH = HOME + "\\local\\bin\\"
os.makedirs(INSTALL_PATH, exist_ok=True)

print(f"\nMouScroll.py コピー先: {INSTALL_PATH}")
copy('MouScroll.pyw', INSTALL_PATH)

# ---------------- 自動起動設定(タスクスケジューラ) ---------------- #
from socket import gethostname
import getpass

# ホスト名取得
HOSTNAME = gethostname()

# 現在のユーザーのSIDを取得
cmd = f'wmic useraccount where name="{getpass.getuser()}" get sid'
output_str = subprocess.run(cmd, capture_output=True, text=True).stdout
SID = output_str.replace("SID", "").replace(" ", "").replace("\n", "")

with open(".\\StartMouScroll.xml", 'r', encoding='utf-16-le') as f:
    text = f.read()

with open(".\\StartMouScroll_cp.xml", 'w+', encoding='utf-16-le') as f:
    body = text.replace("<Author></Author>", f"<Author>{HOSTNAME}</Author>")
    body = body.replace("<UserId></UserId>", f"<UserId>{SID}</UserId>")
    body = body.replace("<Command></Command>", f"<Command>{INSTALL_PATH+'MouScroll.pyw'}</Command>")
    f.write(body)
# プリインストールの PowerShell
PATH_TO_POWERSHELL = "C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe"
# subprocessはPowerShellで実行
os.environ['COMSPEC'] = PATH_TO_POWERSHELL

# 管理者で schtasks の実行
path_to_XML = os.path.dirname(__file__) + "\\StartMouScroll_cp.xml"
cmd = [
    "Start-Process",
    "schtasks.exe",
    "-Verb",
    "RunAs",
    "-WindowStyle",
    "Hidden",
    "-ArgumentList",
    f'"/Create /TN StartMouScroll /XML {path_to_XML}"'  # 新しくプロセス起動してるので絶対パス
]

subprocess.run(cmd, shell=True)
