import os

import openai
from colorama import Fore
from dotenv import load_dotenv

load_dotenv()
openai.api_key = os.getenv("API_KEY")


class ChatGPT:
    exit_command: str = "exit()"
    default_model: str = "gpt-3.5-turbo"

    def __init__(self, summary_length: int = 10):
        self.chat_history: list[dict] = []
        self.chat_summary: str = ""
        self._initial_prompt: str = ""
        self._summary_length: int = summary_length

    @staticmethod
    def _print_api_error_message(e):
        """APIエラーのメッセージを表示する"""

        e = type(e)  # type(e)とすることで、eのサブクラスの型を取得できる
        if isinstance(e, (openai.error.APIError, openai.error.ServiceUnavailableError)):
            print(f"{Fore.RED}OpenAI側でエラーが発生しています。少し待ってから再度試してください。{Fore.RESET}")
            print("サービス稼働状況は https://status.openai.com/ で確認できます。")
        elif e == openai.error.RateLimitError:
            print(f"{Fore.RED}ネットワークに問題があります。設定を見直すか少し待ってから再度試してください。{Fore.RESET}")
        elif e == openai.error.AuthenticationError:
            print(f"{Fore.RED}APIキーまたはトークンが無効もしくは期限切れです。{Fore.RESET}")

    def fetch_gpt_model_list(self) -> list[str]:
        """
        GPTモデルの一覧を取得

        エラーが発生した場合はプログラムを終了させる
        :return: GPTモデルの一覧
        """

        # GPTモデルの一覧を取得
        try:
            model_list = openai.Model.list()
        except openai.error.OpenAIError as e:
            self._print_api_error_message(e)
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

        models_list = self.fetch_gpt_model_list()

        # モデルの選択
        while True:
            # モデルの一覧を表示
            for i, model in enumerate(models_list):
                print(f"{i}: {model}")

            selected_model = input("使用するモデル番号を入力しEnterキーを押してしてください。"
                                   f"何も入力しない場合は '{self.default_model}' が使われます。: ")
            # 何も入力されなかったとき
            if not selected_model:
                return self.default_model
            # 数字以外が入力されたとき
            elif not selected_model.isdigit():
                print(f"{Fore.RED}数字を入力してください。{Fore.RESET}")
            # 選択肢に存在しない番号が入力されたとき
            elif not int(selected_model) in range(len(models_list)):
                print(f"{Fore.RED}その番号は選択肢に存在しません。{Fore.RESET}")
            # 正常な入力
            elif int(selected_model) in range(len(models_list)):
                return models_list[int(selected_model)]

    def _input_user_prompt(self) -> str:
        """
        ユーザーからの入力を受け付ける

        一番最初のプロンプトで終了コマンドが入力された場合は、チャットを終了する。
        2回目以降のプロンプトで終了コマンドが入力された場合は、空文字を返す。
        それ以外の場合は、ユーザーの入力を返す。
        :return: ユーザーのプロンプト及び空文字
        """

        while True:
            # ユーザー入力の受付と履歴への追加
            while True:
                user_prompt = input(f"\n{Fore.CYAN}あなた:{Fore.RESET} ")
                if not user_prompt:
                    print(f"{Fore.YELLOW}プロンプトを入力してください。{Fore.RESET}")
                else:
                    break

            if not self._initial_prompt:
                # いきなり終了コマンドが入力されたときはプログラム終了
                if user_prompt == self.exit_command:
                    exit()
                self._initial_prompt = user_prompt

            if user_prompt == self.exit_command:
                return ""
            else:
                return user_prompt

    def start_chat(self):
        """
        AIアシスタントとユーザーとのチャットを開始し、チャットが終了したら要約を作成する。

        ユーザーからの入力を受け取り、AIアシスタントが応答を生成する。
        ユーザーが exit() と入力すると、チャットは終了する。
        """

        model = self._choice_chat_model()

        print(f"\nAIアシスタントとチャットを始めます。チャットを終了するには {self.exit_command} と入力してください。")
        system_content = input("AIアシスタントに与える役割がある場合は入力してください。"
                               "ない場合はそのままEnterキーを押してください。: ")

        if system_content:
            self.chat_history.append({"role": "system", "content": system_content})

        while True:
            # ユーザーのプロンプトを受け取り履歴に追加する
            user_prompt = self._input_user_prompt()
            if not user_prompt:
                # ユーザーが終了コマンドを入力した場合は空文字が返されチャットを終了する
                break
            self.chat_history.append({"role": "user", "content": user_prompt})

            # GPTによる応答
            completion = openai.ChatCompletion.create(model=model, messages=self.chat_history)
            gpt_answer = completion.choices[0].message.content
            gpt_role = completion.choices[0].message.role

            # 応答の表示と履歴への追加
            print(f"\n{Fore.GREEN}AIアシスタント:{Fore.RESET} {gpt_answer}")
            self.chat_history.append({"role": gpt_role, "content": gpt_answer})

        self._generate_summary()

    def _generate_summary(self):
        """
        チャットの履歴から要約を生成する。

        要約する文字数は self._summary_length で指定する。
        ただし、GPTに文字数を指定して要約を生成させると、指定した文字数よりも多くなる場合がある。
        その場合、要約の文字数を self._summary_length に合わせ、最後に ... を追加する。
        """

        # chat_history の先頭に要約の依頼を追加
        summary_request = {"role": "system",
                           "content": "あなたはユーザーの依頼を要約する役割を担います。"
                                      f"以下のユーザーの依頼を必ず全角{self._summary_length}文字以内で要約してください"}

        # GPTによる要約を取得
        messages = [summary_request, {"role": "user", "content": self._initial_prompt}]
        completion = openai.ChatCompletion.create(model=self.default_model, messages=messages)
        summary = completion.choices[0].message.content

        # 要約を調整
        if len(summary) > self._summary_length:
            self.chat_summary = summary[:self._summary_length] + "..."
        else:
            self.chat_summary = summary
