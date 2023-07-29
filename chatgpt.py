import os

import openai
from colorama import Fore
from dotenv import load_dotenv


def get_gpt_model_list_with_error_handle():
    """
    エラーハンドリングを含むGPTモデルの一覧を取得

    :return: GPTモデルの一覧
    """

    # GPTモデルの一覧を取得
    try:
        model_list = openai.Model.list()
    except (openai.error.APIError, openai.error.ServiceUnavailableError):
        print(f"{Fore.RED}OpenAI側でエラーが発生しています。少し待ってから再度試してください。{Fore.RESET}")
        print("サービス稼働状況は https://status.openai.com/ で確認できます。")
        exit()
    except (openai.error.Timeout, openai.error.APIConnectionError):
        print(f"{Fore.RED}ネットワークに問題があります。設定を見直すか少し待ってから再度試してください。{Fore.RESET}")
        exit()
    except openai.error.AuthenticationError:
        print(f"{Fore.RED}APIキーまたはトークンが無効もしくは期限切れです。{Fore.RESET}")
        exit()
    else:
        models = [model.id for model in model_list.data if "gpt" in model.id]
        models.sort()
        return models


def choice_chat_model():
    """
    GPTモデルの一覧を選択させる

    :return: 選択されたモデルの名称
    """

    models = get_gpt_model_list_with_error_handle()

    # モデルの選択
    while True:
        # モデルの一覧を表示
        for i, model in enumerate(models):
            print(f"{i}: {model}")

        try:
            selected_model = input("使用するモデル番号を入力してください。何も入力しない場合は 'gpt-3.5-turbo' が使われます。: ")
            if not selected_model:
                return "gpt-3.5-turbo"
            return models[int(selected_model)]
        except ValueError:
            print(f"{Fore.RED}数字を入力してください。{Fore.RESET}")
        except IndexError:
            print(f"{Fore.RED}その番号は選択肢に存在しません。{Fore.RESET}")


def chat(model):
    """
    AIアシスタントとユーザーとのチャットを開始する。

    ユーザーからの入力を受け取り、AIアシスタントが応答を生成します。
    ユーザーが 'exit()'と入力すると、チャットは終了します。

    :param model: 使用するGPTモデルの指定
    :return: チャットの履歴。ユーザーとAIアシスタントのロールと発言内容を中身とした辞書のリスト
    """

    print("AIアシスタントとチャットを始めます。チャットを終了するには exit() と入力してください。")
    system_content = input("AIアシスタントに演じて欲しい役割がある場合は入力してください。"
                           "ない場合はそのままEnterキーを押してください。: ")

    # チャットを開始
    chat_history = []
    if system_content:
        chat_history.append({"role": "system", "content": system_content})

    while True:
        # ユーザー入力の受付と履歴への追加
        user_request = input(f"\n{Fore.CYAN}あなた:{Fore.RESET} ")
        if user_request == "exit()":
            break
        chat_history.append({"role": "user", "content": user_request})

        # GPTによる応答
        completion = openai.ChatCompletion.create(model=model, messages=chat_history)
        gpt_answer = completion.choices[0].message.content
        gpt_role = completion.choices[0].message.role
        chat_history.append({"role": gpt_role, "content": gpt_answer})
        print(f"\n{Fore.GREEN}AIアシスタント:{Fore.RESET} {gpt_answer}")

    return chat_history


if __name__ == "__main__":
    # 初期設定
    load_dotenv()
    api_key = os.getenv("API_KEY")
    openai.api_key = api_key

    choice_model = choice_chat_model()
    chat_hist = chat(choice_model)
    print(chat_hist)
