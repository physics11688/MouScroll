from pynput import mouse, keyboard  # マウスとキーボードの操作を拾う
from win32clipboard import (
    OpenClipboard,
    GetClipboardData,
    CloseClipboard,
    CF_UNICODETEXT,
)  # クリップボード操作のため
from subprocess import run
import pyautogui as pg  # キーボード操作用。主にホットキー。
import pygetwindow as gw  # ActiveWindows取得とかに使う
from time import time
from pathlib import Path
from os.path import expanduser
from os import getlogin

# -------------------- キーボードの操作 -------------------- #

# -------------------- マウスの操作  -------------------- #
pg.FAILSAFE = False  # 画面端でも例外吐かないようにする
pre_time = 0  # スクロール時間間隔用
TIME_IV = 0.027  # (要調整1) 処理を飛ばすスクロール時間間隔
POINTER_SPEED = 12  # (要調整2)
CNT_SPEED = 8  # (要調整3)  0-10

# ホットスポット領域の境界座標
LEFT_RANGE = int(pg.size().width * 5 / 100)  # 画面左側5%
RIGHT_RANGE = int(pg.size().width * 95 / 100)  # 画面右側5%
TOP_RANGE = int(pg.size().height * 5 / 100)  # 画面上側5%
BOTTOM_RANGE = int(pg.size().height * 95 / 100)  # 画面下側5%
X_MIDDLE_RANGE = {
    "LEFT": int(pg.size().width * 40 / 100),
    "RIGHT": int(pg.size().width * 60 / 100),
}
Y_MIDDLE_RANGE = {
    "TOP": int(pg.size().height * 35 / 100),
    "BOTTOM": int(pg.size().height * 65 / 100),
}

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

# みんな大好き MicrosoftEdgeのパス
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


def rewrite(keyword):
    # 新しいファイルエクスプローラの仕様対策
    if " とその他" in keyword:
        keyword = keyword[: keyword.find(" とその他")]
    elif " - エクスプローラー" in keyword:
        keyword = keyword[: keyword.find(" - エクスプローラー")]
    # ドキュメントを選択してるとウィンドウタイトルが日本語
    keyword = (
        keyword.replace("ミュージック", "Music")
        .replace("ビデオ", "Videos")
        .replace("キャプチャ", "Captures")
        .replace("ピクチャ", "Pictures")
        .replace("カメラロール", "Camera Roll")
        .replace("スクリーンショット", "Screenshots")
        .replace("ドキュメント", "Documents")
        .replace("デスクトップ", "Desktop")
        .replace("ダウンロード", "Downloads")
        .replace("ライブラリ\\", "")  # ライブラリを直に
        .replace("ライブラリ", expanduser("~"))  # ライブラリ内
    )

    # ホームディレクトリ対策(Microsoftアカウント)
    if (
        keyword
        in run(["net", "user", getlogin()], capture_output=True, text=True).stdout
    ):
        keyword = expanduser("~")

    if keyword in [
        "Music",
        "Videos",
        "Captures",
        "Pictures",
        "Camera Roll",
        "Screenshots",
        "Documents",
        "Desktop",
        "Downloads",
    ]:
        keyword = expanduser("~") + "\\" + keyword
        if not Path(keyword).exists():
            keyword = expanduser("~") + "\\OneDrive\\" + keyword

    # デスクトップ使用中
    if keyword == "Program Manager":
        keyword = expanduser("~") + "\\" + "Desktop"
        if not Path(keyword).exists():
            keyword = expanduser("~") + "\\OneDrive\\" + "Desktop"

    return keyword


def on_scroll(x, y, dx, dy) -> None:
    """マウスをスクロールした時呼ばれる関数

    Args:
        x (int): マウスポインタのx座標
        y (int): マウスポインタのy座標
        dx (int): x軸方向のスクロール
        dy (int): y軸方向のスクロール
    """

    global wheel
    global pre_time

    if time() - pre_time < TIME_IV:  # 前回の処理時間との差が短すぎるなら
        pre_time = time()  # 時間だけ格納し直して
        return  # 処理を飛ばす

    # 下にスクロール
    if dy < 0:
        # 仮想デスクトップ切り替え
        if (x > RIGHT_RANGE) and (y < TOP_RANGE):  # カーソルが ➚
            pg.hotkey("ctrl", "win", "left")  # 次の仮想デスクトップへ

        # タスクの切り替え
        elif (x < LEFT_RANGE) and (y < TOP_RANGE):  # カーソルが ↖
            # 既にタスク切換えウィンドウが出てるなら
            if "タスクの切り替え" in gw.getAllTitles():
                pg.keyDown("left")  # 右キーでタスク選択を移動
            else:  # タスク切換えウィンドウが出てないなら
                pg.hotkey("ctrl", "alt", "tab")  # タスク切換え
        # デスクトップの表示
        elif (x < LEFT_RANGE) and (y > BOTTOM_RANGE):  # カーソルが ↙
            # アクティブウィンドウがないなら
            current_HWND = gw.getActiveWindow()
            if (current_HWND is None) or (current_HWND.title == ""):
                pass
            else:  # あるならデスクトップ表示
                pg.hotkey("win", "d")  # 同時押し
        # 選択文字列で検索
        elif (x > RIGHT_RANGE) and (y > BOTTOM_RANGE):  # カーソルが➘
            pg.hotkey("ctrl", "c")
            args = "microsoft-edge:https://www.google.co.jp/search?q=" + get_Clip()
            run([PATH_TO_EDGE, args])
        elif (X_MIDDLE_RANGE["LEFT"] < x < X_MIDDLE_RANGE["RIGHT"]) and (
            y > BOTTOM_RANGE
        ):
            pg.press("volumemute")
        elif (x < LEFT_RANGE) and (
            Y_MIDDLE_RANGE["TOP"] < y < Y_MIDDLE_RANGE["BOTTOM"]
        ):
            # pg.press("browserback")
            pg.hotkey("alt", "tab")  # タスク切換え
        elif (x > RIGHT_RANGE) and (
            Y_MIDDLE_RANGE["TOP"] < y < Y_MIDDLE_RANGE["BOTTOM"]
        ):
            path = rewrite(gw.getActiveWindow().title)
            if Path(path).exists():
                run(
                    [
                        "code",
                        rewrite(f"{path}"),
                    ],
                    shell=True,
                )

    # 上にスクロール
    else:
        # 仮想デスクトップ切り替え
        if (x > RIGHT_RANGE) and (y < TOP_RANGE):  # カーソルが ➚
            pg.hotkey("ctrl", "win", "right")  # 前の仮想デスクトップへ

        # タスクの切り替え
        elif (x < LEFT_RANGE) and (y < TOP_RANGE):  # カーソルが ↖
            # 既にタスク切換えウィンドウが出てるなら
            if "タスクの切り替え" in gw.getAllTitles():
                #  左キーでタスク選択を移動
                pg.keyDown("right")
            else:  # タスク切換えウィンドウが出てないなら
                pg.hotkey("ctrl", "alt", "tab")  # タスク切換え

        # アプリの表示
        elif (x < LEFT_RANGE) and (y > BOTTOM_RANGE):  # カーソルが ↙
            # アクティブウィンドウがないなら
            current_HWND = gw.getActiveWindow()
            if (current_HWND is None) or (current_HWND.title == ""):
                pass
            else:  # あるならデスクトップ表示
                pg.hotkey("win", "d")  # 同時押し
        # 選択文字列でプライベート検索
        elif (x > RIGHT_RANGE) and (y > BOTTOM_RANGE):  # カーソルが➘
            pg.hotkey("ctrl", "c")
            args = "microsoft-edge:https://www.google.co.jp/search?q=" + get_Clip()
            run([PATH_TO_EDGE, "--inprivate", args])

        elif (X_MIDDLE_RANGE["LEFT"] < x < X_MIDDLE_RANGE["RIGHT"]) and (y < TOP_RANGE):
            pg.hotkey("win", "shift", "s")  # スクショ
        elif (x < LEFT_RANGE) and (
            Y_MIDDLE_RANGE["TOP"] < y < Y_MIDDLE_RANGE["BOTTOM"]
        ):
            # pg.press("browserforward")
            pg.hotkey("alt", "tab")  # タスク切換え
        elif (x > RIGHT_RANGE) and (
            Y_MIDDLE_RANGE["TOP"] < y < Y_MIDDLE_RANGE["BOTTOM"]
        ):
            path = rewrite(gw.getActiveWindow().title)
            if Path(path).exists():
                run(
                    ["wt.exe", "-d", rewrite(f"{path}"), "pwsh.exe"],
                    shell=True,
                )

    pre_time = time()
    return


def on_click(x, y, button, pressed) -> None:
    """カーソル ↖ でタスク切換え中にクリックでアプリ選択"""
    isUpperLeft = (x < LEFT_RANGE) and (y < TOP_RANGE)
    if isUpperLeft and ("タスクの切り替え" in gw.getAllTitles()):  # カーソルが ↖
        pg.keyDown("enter")
    return


x_pointer = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]


def move(x, y):
    # if y == 0:
    #     raise ValueError("error!")

    global x_pointer
    x_pointer = x_pointer[1:] + [x]
    if (LEFT_RANGE < x < RIGHT_RANGE) or not (TOP_RANGE < y < BOTTOM_RANGE):
        return

    s_list = [x_pointer[i + 1] - x_pointer[i] for i in range(len(x_pointer) - 1)]

    if (any((s >= 0 for s in s_list))) and len(
        [s for s in s_list if s > POINTER_SPEED]
    ) > CNT_SPEED:
        x_pointer = [x] * 10
        pg.hotkey("ctrl", "win", "right")

        return
    elif (any((s <= 0 for s in s_list))) and len(
        [s for s in s_list if s < -POINTER_SPEED]
    ) > CNT_SPEED:
        x_pointer = [x] * 10
        pg.hotkey("ctrl", "win", "left")


# りっすん
with mouse.Listener(on_move=move, on_scroll=on_scroll, on_click=on_click) as listener:
    try:
        listener.join()  # ずーとループしてリッスンしてる
    except Exception as e:
        # 例外吐いたら一応知らせる
        pg.alert(
            text=f"例外が発生したため、MouScroll.pyを終了しました.\n{e}",
            title="MouScroll",
            button="OK",
        )
