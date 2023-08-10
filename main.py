from chatgpt import ChatGPT
from output_excel import output_excel, excel_path, is_chat_history_open

if is_chat_history_open() and excel_path.exists():
    print(f"{excel_path.name} が開かれています。閉じてから再度実行してください。")
else:
    gpt = ChatGPT()
    gpt.start_chat()
    output_excel(gpt)
