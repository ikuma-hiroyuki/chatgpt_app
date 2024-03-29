from chatgpt import ChatGPT
from output_excel import output_excel, excel_path, is_open_output_excel

if is_open_output_excel() and excel_path.exists():
    print(f"{excel_path.name} が開かれています。閉じてから再度実行してください。")
else:
    gpt = ChatGPT()
    gpt.chat_runner()
    output_excel(gpt)
