from pynput import mouse   # マウスの操作を拾う
from win32con import MB_OK, MB_ICONWARNING, NULL  # Win32API使う。pywin32の一部
from win32clipboard import OpenClipboard, GetClipboardData, CloseClipboard  # クリップボード操作のため
from subprocess import Popen
from plyer import notification  # Desktop notification

import pyautogui as pg     # キーボード操作用。主にホットキー。
import pygetwindow as gw   # ActiveWindows取得とかに使う


pg.FAILSAFE = False  # 画面端でも例外吐かないようにする
wheel = 0  # 検索用のホイールカウンタ

# ホットスポット領域の境界座標
L_range = 100
R_range = 1770
U_range = 1030
D_range = 50


def get_Clip() -> str:
    """クリップボードの内容を取得する関数"""
    try:
        OpenClipboard()
    except:
        return ""  # オープン出来なかったら空文字を返す

    data = GetClipboardData()  # 内容をゲット
    CloseClipboard()

    # dataが文字列なら中身をreturn
    return data if isinstance(data, str) else ""


def on_scroll(x, y, dx, dy) -> None:
    """マウスをスクロールした時呼ばれる関数"""
    global wheel
    #print('Scrolled {0} at {1}'.format('down' if dy < 0 else 'up',(x, y)))

    # 下にスクロール
    if (dy < 0):
        # 仮想デスクトップ切り替え
        if (x > R_range) and (y < D_range):  # カーソルが ➚
            pg.hotkey('ctrl', 'win', 'left')  # 次の仮想デスクトップへ
            return

        # タスクの切り替え
        if (x < L_range) and (y < D_range):  # カーソルが ↖
            # 既にタスク切換えウィンドウが出てるなら
            if "タスクの切り替え" in gw.getAllTitles():
                #  右キーでタスク選択を移動
                pg.keyDown('right')
            else:  # タスク切換えウィンドウが出てないなら
                pg.hotkey('ctrl', 'alt', 'tab')  # タスク切換え
            return

        # デスクトップの表示
        if (x < L_range) and (y > U_range):  # カーソルが ↙
            # アクティブウィンドウがないならreturn
            if gw.getActiveWindow().title == "":
                return
            else:  # あるならデスクトップ表示
                pg.hotkey('win', 'd')  # 同時押し
            return

        # 選択文字列で検索
        if (x > R_range) and (y > U_range):  # カーソルが➘
            # 選択中の文字列をクリップボードへコピー
            pg.hotkey('ctrl', 'c')

            edge = "C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe"
            cmd = "microsoft-edge:https://www.google.co.jp/search?q=" + get_Clip()

            if wheel >= 4:
                Popen([edge, cmd])
                wheel = 0
            else:
                wheel += 1
            return

    # 上にスクロール
    else:
        # 仮想デスクトップ切り替え
        if (x > R_range) and (y < D_range):   # カーソルが ➚
            pg.hotkey('ctrl', 'win', 'right')  # 前の仮想デスクトップへ
            return

        # タスクの切り替え
        if (x < L_range) and (y < D_range):  # カーソルが ↖
            # 既にタスク切換えウィンドウが出てるなら
            if "タスクの切り替え" in gw.getAllTitles():
                #  左キーでタスク選択を移動
                pg.keyDown('left')
            else:   # タスク切換えウィンドウが出てないなら
                pg.hotkey('ctrl', 'alt', 'tab')  # タスク切換え
            return

        # アプリの表示
        if (x < L_range) and (y > U_range):  # カーソルが ↙
            # アクティブウィンドウないならタスク表示
            if gw.getActiveWindow().title == "":
                pg.hotkey('win', 'd')  # 同時押し
            else:
                return
            return

        # 選択文字列で検索
        if (x > R_range) and (y > U_range):  # カーソルが➘
            pg.hotkey('ctrl', 'c')
            edge = "C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe"
            cmd = "microsoft-edge:https://www.google.co.jp/search?q=" + get_Clip()

            if wheel >= 4:
                Popen([edge, cmd])
                wheel = 0
            else:
                wheel += 1
            return


def on_click(x, y, button, pressed) -> None:
    """カーソル ↖ でタスク切換え中にクリックでアプリ選択"""
    if (x < L_range) and (y < D_range) and ("タスクの切り替え" in gw.getAllTitles()):  # カーソルが ↖
        pg.keyDown('enter')
    return


# りっすん
with mouse.Listener(on_scroll=on_scroll, on_click=on_click) as listener:
    try:
        listener.join()  # ずーとループしてリッスンしてる
    except Exception as e:
        # 例外はいたらnotificationで知らせる
        notification.notify('MouScroll', '例外が発生したため、MouScroll.pyを終了しました。')
