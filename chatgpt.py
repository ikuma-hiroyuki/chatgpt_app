import os

import openai
from colorama import Fore
from dotenv import load_dotenv

from output_excel import is_chat_history_open, excel_path, output_excel

# 初期設定
load_dotenv()
openai.api_key = os.getenv("API_KEY")


class ChatGPT:
    def __init__(self, summary_length: int = 10):
        self.chat_history: list[dict] = []
        self.chat_summary: str = ""
        self.initial_prompt: str = ""
        self.summary_length: int = summary_length

    @staticmethod
    def _api_error_message(e):
        """APIエラーのメッセージを表示する"""

        if e in [openai.error.APIError, openai.error.ServiceUnavailableError]:
            print(f"{Fore.RED}OpenAI側でエラーが発生しています。少し待ってから再度試してください。{Fore.RESET}")
            print("サービス稼働状況は https://status.openai.com/ で確認できます。")
        elif e in [openai.error.Timeout, openai.error.APIConnectionError]:
            print(f"{Fore.RED}ネットワークに問題があります。設定を見直すか少し待ってから再度試してください。{Fore.RESET}")
        elif e == openai.error.AuthenticationError:
            print(f"{Fore.RED}APIキーまたはトークンが無効もしくは期限切れです。{Fore.RESET}")

    def fetch_gpt_model_list(self) -> list[str]:
        """
        GPTモデルの一覧を取得

        :return: GPTモデルの一覧
        """

        # GPTモデルの一覧を取得
        try:
            model_list = openai.Model.list()
        except openai.error.OpenAIError as e:
            self._api_error_message(type(e))  # type(e)とすることで、eのサブクラスの型を取得できる
            exit()
        else:
            models = [model.id for model in model_list.data if "gpt" in model.id]
            models.sort()
            return models

    def _choice_chat_model(self) -> str:
        """
        GPTモデルの一覧を選択させる

        :return: 選択されたモデルの名称
        """

        default_model = "gpt-3.5-turbo"
        models_list = self.fetch_gpt_model_list()

        # モデルの選択
        while True:
            # モデルの一覧を表示
            for i, model in enumerate(models_list):
                print(f"{i}: {model}")

            selected_model = input("使用するモデル番号を入力しEnterキーを押してしてください。"
                                   f"何も入力しない場合は '{default_model}' が使われます。: ")
            # 何も入力されなかったとき
            if not selected_model:
                return default_model
            # 数字以外が入力されたとき
            elif not selected_model.isdigit():
                print(f"{Fore.RED}数字を入力してください。{Fore.RESET}")
            # 選択肢に存在しない番号が入力されたとき
            elif not int(selected_model) in range(len(models_list)):
                print(f"{Fore.RED}その番号は選択肢に存在しません。{Fore.RESET}")
            # 正常な入力
            elif int(selected_model) in range(len(models_list)):
                return models_list[int(selected_model)]

    def _start_chat(self):
        """
        AIアシスタントとユーザーとのチャットを開始し、チャットが終了したら要約を作成する。

        ユーザーからの入力を受け取り、AIアシスタントが応答を生成します。
        ユーザーが 'exit()'と入力すると、チャットは終了します。
        """

        exit_command = "exit()"
        model = self._choice_chat_model()

        print(f"\nAIアシスタントとチャットを始めます。チャットを終了するには {exit_command} と入力してください。")
        system_content = input("AIアシスタントに演じて欲しい役割がある場合は入力してください。"
                               "ない場合はそのままEnterキーを押してください。: ")

        # チャットを開始
        if system_content:
            self.chat_history.append({"role": "system", "content": system_content})

        while True:
            # ユーザー入力の受付と履歴への追加
            while True:
                user_request = input(f"\n{Fore.CYAN}あなた:{Fore.RESET} ")
                if not user_request:
                    print(f"{Fore.YELLOW}プロンプトを入力してください。{Fore.RESET}")
                else:
                    break

            if not self.initial_prompt:
                self.initial_prompt = user_request

            if self.initial_prompt and user_request == exit_command:
                break
            elif user_request == exit_command:
                exit()

            self.chat_history.append({"role": "user", "content": user_request})

            # GPTによる応答
            completion = openai.ChatCompletion.create(model=model, messages=self.chat_history)
            gpt_answer = completion.choices[0].message.content
            gpt_role = completion.choices[0].message.role

            # 応答の表示と履歴への追加
            print(f"\n{Fore.GREEN}AIアシスタント:{Fore.RESET} {gpt_answer}")
            self.chat_history.append({"role": gpt_role, "content": gpt_answer})

        self._generate_summary()

    def _generate_summary(self):
        """ チャットの履歴から要約を生成する。 """

        # chat_history の先頭に要約の依頼を追加
        summary_request = {"role": "system",
                           "content": f"あなたはユーザーの依頼を要約する役割を担います。以下のユーザーの依頼を必ず全角{self.summary_length}文字以内で要約してください"}

        # GPTによる要約を取得
        messages = [summary_request, {"role": "user", "content": self.initial_prompt}]
        completion = openai.ChatCompletion.create(model="gpt-3.5-turbo", messages=messages)
        self.chat_summary = completion.choices[0].message.content

        # 要約を調整
        if len(self.chat_summary) > self.summary_length:
            self.chat_summary = self.chat_summary[:self.summary_length] + "..."

    def run(self):
        if is_chat_history_open():
            print(f"{excel_path.name} が開かれています。閉じてから再度実行してください。")
        else:
            self._start_chat()
            output_excel(self)


if __name__ == "__main__":
    chat = ChatGPT()
    chat.run()
