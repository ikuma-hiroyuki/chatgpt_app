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


def output_excel(chat):
    """
    chat_history.xlsx にチャットの履歴を書き込む。
    """

    # ワークブックの読み込み
    is_new = False
    if excel_path.exists():
        workbook = openpyxl.load_workbook(excel_path)
    else:
        workbook = openpyxl.Workbook()
        is_new = True

    # ワークシートの作成
    if is_new:
        worksheet = workbook.active
        worksheet.title = chat.chat_summary
    else:
        worksheet = workbook.create_sheet(title=chat.chat_summary)  # 同じ名前がある場合、末尾に数字が付与される
        workbook.active = worksheet

    # ヘッダーの書き込み
    worksheet["A1"], worksheet["B1"] = "ロール", "発言内容"

    # データの書き込み
    for i, content in enumerate(chat.chat_history, 2):
        worksheet[f"A{i}"], worksheet[f"B{i}"] = content["role"], content["content"]

    # ワークブックの保存
    workbook.save(excel_path)
    workbook.close()
