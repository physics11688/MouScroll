from platform import system
import os
import re
import subprocess


def YorN() -> bool:
    user_input = input("Yes or No >> ")

    Yes = re.search(r"/yes/i", user_input)
    Y = re.search(r"[Yy]", user_input)
    if Yes is Y is None:
        return False
    else:
        return True


# ディレクトリの移動
os.chdir(os.path.dirname(__file__))

# OSの確認
if system() != 'Windows':
    print("MouScrollは Windows専用のプログラムです")
    print("プログラムを終了します")
    exit(2)

print("MouScroll をアンインストールしますか?")
if not YorN():
    print("プログラムを終了します")
    exit(1)

HOME = os.path.expanduser("~")  # ホームディレクトリのパス取得

if " " in HOME:
    print("\nホームディレクトリのPATHに 空白 が含まれています.")
    print("出直してきやがれ")
    exit(0)

INSTALL_PATH = HOME + "\\local\\bin\\auto_auth.pyw"

if os.path.isfile(INSTALL_PATH):
    os.remove(INSTALL_PATH)

print(INSTALL_PATH + " を削除しました.")

# ----------------- 自動起動設定の削除 ----------------- #

# Windows: タスクスケジューラの設定
# プリインストールの PowerShell
PATH_TO_POWERSHELL = "C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe"
# subprocessはPowerShellで実行
os.environ['COMSPEC'] = PATH_TO_POWERSHELL

# 管理者で schtasks の実行
cmd = ["Start-Process", "schtasks.exe", "-Verb", "RunAs", "-ArgumentList", '"/Delete /TN StartMouScroll"']
subprocess.run(cmd, shell=True)
print("タスク: StartMouScroll を削除しました")

print("アンインストールが正常に終了しました")
