from pathlib import Path

import openpyxl

excel_path = Path("chat_history.xlsx")


def is_chat_history_open() -> bool:
    """excel_pathが開かれているかどうかを返す"""
    try:
        with excel_path.open("r+b"):
            pass
        return False
    except IOError:
        return True


def write_excel(chat_history: list[dict], worksheet_title: str):
    """
    chat_history.xlsx にチャットの履歴を書き込む。

    :param chat_history: チャットの履歴。ユーザーとAIアシスタントのロールと発言内容を中身とした辞書のリスト
    :param worksheet_title: ワークシートのタイトル
    """

    # ワークブックの読み込み
    is_new = False
    if excel_path.exists():
        wb = openpyxl.load_workbook(excel_path)
    else:
        wb = openpyxl.Workbook()
        is_new = True

    # ワークシートの作成
    if is_new:
        ws = wb.active
        ws.title = worksheet_title
    else:
        ws = wb.create_sheet(title=worksheet_title)
        wb.active = ws

    # ヘッダーの書き込み
    ws["A1"] = "ロール"
    ws["B1"] = "発言内容"

    # データの書き込み
    for i, content in enumerate(chat_history):
        ws[f"A{i + 2}"] = content["role"]
        ws[f"B{i + 2}"] = content["content"]

    # ワークブックの保存
    wb.save(excel_path)
    wb.close()
