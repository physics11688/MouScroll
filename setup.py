import subprocess
import re
from platform import system
import sys
from pathlib import Path


def color_print(color: int, text: str, is_bold=False, end="\n") -> None:
    """色付きのprint関数

    Args:
        color: 色番号

        text: prints帯文字列

        is_bold: ボールドにしたきゃ True

        end: 改行したくないなら end=""
    """
    if is_bold:
        text = "\033[1m{:10s}\033[0m".format(text)
    if end == "":
        print("\033[{0}m{1}\033[0m".format(color, text), end="")
    else:
        print("\033[{0}m{1}\033[0m".format(color, text))


def YorN() -> bool:
    user_input = input("Yes or No >> ")

    Yes = re.search(r"/yes/i", user_input)
    Y = re.search(r"[Yy]", user_input)
    if Yes is Y is None:
        return False
    else:
        return True


msg = """これはMouScroll.appのインストールスクリプトです。
以下の処理を行います。

1. pipを使って pynput, pywin32, pyautogui, pygetwindow, plyer のインストール

2. ホームディレクトリ直下に local\\bin\\MouScroll フォルダを作り、スクリプト MouScroll.py を突っ込む



"""

print(msg)

color_print(35, "セットアップを実行しますか？\n")
if not YorN():
    print("プログラムを終了します")
    sys.exit()

# OSの確認
if system() != 'Windows':
    print("MouScrollは Windowsのみ使用可能です")
    print("プログラムを終了します")
    sys.exit()


try:
    subprocess.run(["pip", "install", "-U", "pynput", "pywin32", "pyautogui", "pygetwindow", "plyer"])
except:
    print("モジュールのインストールに失敗しました")
    print("プログラムを終了します")

home = Path.home()
install_dir_path = Path(home / "local" / "bin" / "MouScroll")
# インストールディレクトリの作成
install_dir_path.mkdir(parents=True, exist_ok=True)

for path in Path(__file__).parent.iterdir():
    if path.is_file():
        name = path.name
        new_path = install_dir_path / name
        path.rename(new_path)
