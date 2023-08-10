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


def trim_invalid_chars(sheet_name: str) -> str:
    """エクセルのシート名で使えない文字列を削除"""
    invalid_chars = [':', '\\', '/', '?', '*', '[', ']']
    for char in invalid_chars:
        sheet_name = sheet_name.replace(char, '')
    return sheet_name


def output_excel(gpt):
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

    sheet_name = trim_invalid_chars(gpt.chat_summary)

    # ワークシートの作成
    if is_new:
        worksheet = workbook.active
        worksheet.title = sheet_name
    else:
        # 同じ名前がある場合、末尾に数字が付与される
        worksheet = workbook.create_sheet(title=sheet_name)
        workbook.active = worksheet

    # ヘッダーの書き込み
    worksheet["A1"], worksheet["B1"] = "ロール", "発言内容"

    # データの書き込み
    for i, content in enumerate(gpt.chat_history, 2):
        worksheet[f"A{i}"], worksheet[f"B{i}"] = content["role"], content["content"]

    # ワークブックの保存
    workbook.save(excel_path)
    workbook.close()
