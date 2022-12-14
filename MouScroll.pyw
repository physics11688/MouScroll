from pynput import mouse, keyboard  # マウスとキーボードの操作を拾う
from win32clipboard import OpenClipboard, GetClipboardData, CloseClipboard, CF_UNICODETEXT  # クリップボード操作のため
from subprocess import run
import pyautogui as pg  # キーボード操作用。主にホットキー。
import pygetwindow as gw  # ActiveWindows取得とかに使う
from time import time

# -------------------- キーボードの操作 -------------------- #

# -------------------- マウスの操作  -------------------- #
pg.FAILSAFE = False  # 画面端でも例外吐かないようにする
pre_time = 0  # スクロール時間間隔用
TIME_IV = 0.027  # 処理を飛ばす飛ばすスクロール時間間隔

# ホットスポット領域の境界座標
LEFT_RANGE = int(pg.size().width * 5 / 100)  # 画面左側5%
RIGHT_RANGE = int(pg.size().width * 95 / 100)  # 画面右側5%
TOP_RANGE = int(pg.size().height * 5 / 100)  # 画面上側5%
BOTTOM_RANGE = int(pg.size().height * 95 / 100)  # 画面下側5%

# https://pyautogui.readthedocs.io/en/latest/mouse.html
# 0,0       X increases -->
# +---------------------------+
# |                           | Y increases
# |                           |     |
# |   1920 x 1080 screen      |     |
# |                           |     V
# |                           |
# |                           |
# +---------------------------+ 1919, 1079

PATH_TO_EDGE = "C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe"


def get_Clip() -> str:
    """クリップボードの内容を取得する関数"""
    try:
        OpenClipboard()
        data = GetClipboardData(CF_UNICODETEXT)  # 内容を取得
        CloseClipboard()
    except:
        return ""  # オープン出来なかったら空文字を返す

    # dataが文字列なら中身をreturn
    return data if isinstance(data, str) else ""


def on_scroll(x, y, dx, dy) -> None:
    """ マウスをスクロールした時呼ばれる関数

    Args:
        x (_type_): マウスポインタのx座標
        y (_type_): マウスポインタのy座標
        dx (_type_): x軸方向のスクロール
        dy (_type_): y軸方向のスクロール
    """

    global wheel
    global pre_time

    if time() - pre_time < TIME_IV:
        pre_time = time()
        return

    # 下にスクロール
    if (dy < 0):
        # 仮想デスクトップ切り替え
        if (x > RIGHT_RANGE) and (y < TOP_RANGE):  # カーソルが ➚
            pg.hotkey('ctrl', 'win', 'left')  # 次の仮想デスクトップへ

        # タスクの切り替え
        elif (x < LEFT_RANGE) and (y < TOP_RANGE):  # カーソルが ↖
            # 既にタスク切換えウィンドウが出てるなら
            if "タスクの切り替え" in gw.getAllTitles():
                pg.keyDown('left')  # 右キーでタスク選択を移動
            else:  # タスク切換えウィンドウが出てないなら
                pg.hotkey('ctrl', 'alt', 'tab')  # タスク切換え
        # デスクトップの表示
        elif (x < LEFT_RANGE) and (y > BOTTOM_RANGE):  # カーソルが ↙
            # アクティブウィンドウがないなら
            current_HWND = gw.getActiveWindow()
            if (current_HWND is None) or (current_HWND.title == ""):
                pass
            else:  # あるならデスクトップ表示
                pg.hotkey('win', 'd')  # 同時押し
        # 選択文字列で検索
        elif (x > RIGHT_RANGE) and (y > BOTTOM_RANGE):  # カーソルが➘

            pg.hotkey('ctrl', 'c')
            args = "microsoft-edge:https://www.google.co.jp/search?q=" + get_Clip()
            run([PATH_TO_EDGE, args])

    # 上にスクロール
    else:
        # 仮想デスクトップ切り替え
        if (x > RIGHT_RANGE) and (y < TOP_RANGE):  # カーソルが ➚
            pg.hotkey('ctrl', 'win', 'right')  # 前の仮想デスクトップへ

        # タスクの切り替え
        elif (x < LEFT_RANGE) and (y < TOP_RANGE):  # カーソルが ↖
            # 既にタスク切換えウィンドウが出てるなら
            if "タスクの切り替え" in gw.getAllTitles():
                #  左キーでタスク選択を移動
                pg.keyDown('right')
            else:  # タスク切換えウィンドウが出てないなら
                pg.hotkey('ctrl', 'alt', 'tab')  # タスク切換え

        # アプリの表示
        elif (x < LEFT_RANGE) and (y > BOTTOM_RANGE):  # カーソルが ↙
            # アクティブウィンドウないならタスク表示
            if gw.getActiveWindow().title == "":
                pg.hotkey('win', 'd')  # 同時押し

        # 選択文字列でプライベート検索
        elif (x > RIGHT_RANGE) and (y > BOTTOM_RANGE):  # カーソルが➘
            pg.hotkey('ctrl', 'c')
            args = "microsoft-edge:https://www.google.co.jp/search?q=" + get_Clip()
            run([PATH_TO_EDGE, "--inprivate", args])

    pre_time = time()
    return


def on_click(x, y, button, pressed) -> None:
    """ カーソル ↖ でタスク切換え中にクリックでアプリ選択 """
    if (x < LEFT_RANGE) and (y < TOP_RANGE) and ("タスクの切り替え" in gw.getAllTitles()):  # カーソルが ↖
        pg.keyDown('enter')
    return


# りっすん
with mouse.Listener(on_scroll=on_scroll, on_click=on_click) as listener:
    try:
        listener.join()  # ずーとループしてリッスンしてる
    except Exception as e:
        # 例外吐いたら一応知らせる
        pg.alert(text=f"例外が発生したため、MouScroll.pyを終了しました.\n{e}", title="MouScroll", button="OK")
