import os

import openai
from colorama import Fore
from dotenv import load_dotenv

load_dotenv()
openai.api_key = os.getenv("API_KEY")


class ChatGPT:
    EXIT_COMMAND: str = "exit()"
    DEFAULT_MODEL: str = "gpt-3.5-turbo"

    def __init__(self, summary_length: int = 10):
        self.gpt_model = ""
        self.chat_history: list[dict] = []
        self.chat_summary: str = ""
        self._initial_prompt: str = ""
        self._summary_length: int = summary_length

    @staticmethod
    def _print_error_message(message):
        """エラーメッセージを表示する"""
        print(f"{Fore.RED}{message}{Fore.RESET}")

    @staticmethod
    def _throw_api_errors(e):
        """APIエラーを振り分ける"""

        e = type(e)  # type(e)とすることで、eのサブクラスの型を取得できる
        if isinstance(e, (openai.error.APIError, openai.error.ServiceUnavailableError)):
            ChatGPT._print_error_message("OpenAI側でエラーが発生しています。少し待ってから再度試してください。")
            print("サービス稼働状況は https://status.openai.com/ で確認できます。")
        elif isinstance(e, openai.error.RateLimitError):
            ChatGPT._print_error_message("ネットワークに問題があります。設定を見直すか少し待ってから再度試してください。")
        elif isinstance(e, openai.error.AuthenticationError):
            ChatGPT._print_error_message("APIキーまたはトークンが無効もしくは期限切れです。")

    @staticmethod
    def fetch_gpt_model_list() -> list[str] | None:
        """
        GPTモデルの一覧を取得する。APIエラーでも出るが取得できないときはNoneを返す。
        :return: GPTモデルの一覧もしくはNone
        """

        try:
            model_list = openai.Model.list()
        except openai.error.OpenAIError as e:
            ChatGPT._throw_api_errors(e)
            return None
        else:
            models = [model.id for model in model_list.data if "gpt" in model.id]
            models.sort()
            return models

    def _choice_chat_model(self) -> str | None:
        """
        GPTモデルの一覧を選択させる。APIエラーでモデルが取得できないときはNoneを返す。
        :return: 選択されたモデルの名称もしくはNone
        """

        models_list = self.fetch_gpt_model_list()
        if not models_list:
            return None

        while True:
            # モデルの一覧を表示
            print("\nAIとのチャットに使用するモデルの番号を入力しEnterキーを押してしてください。")
            for i, model in enumerate(models_list):
                print(f"{i}: {model}")
            selected_model = input(f"何も入力しない場合は{Fore.GREEN} {self.DEFAULT_MODEL} {Fore.RESET}が使われます。: ")

            # 何も入力されなかったとき
            if not selected_model:
                return self.DEFAULT_MODEL
            # 数字以外が入力されたとき
            if not selected_model.isdigit():
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
        :return: ユーザーのプロンプト
        """

        while True:
            while True:
                user_prompt = input(f"{Fore.CYAN}あなた:{Fore.RESET} ")
                if not user_prompt:
                    print(f"{Fore.YELLOW}プロンプトを入力してください。{Fore.RESET}")
                else:
                    break

            if not self._initial_prompt:
                self._initial_prompt = user_prompt

            return user_prompt

    def _give_role_to_gpt(self):
        """ AIアシスタントに与える役割を入力させる """
        print(f"AIアシスタントとチャットを始めます。")
        system_content = input("AIアシスタントに与える役割がある場合は入力してください。"
                               "ない場合はそのままEnterキーを押してください。: ")
        if system_content:
            self.chat_history.append({"role": "system", "content": system_content})

    def _fetch_gpt_answer(self):
        """
        GPTモデルにユーザーのプロンプトを与えて応答を生成させ、チャット履歴に追加するとともに、応答を返却する。
        :returns: AIアシスタントの応答と役割
        """

        completion = openai.ChatCompletion.create(model=self.gpt_model, messages=self.chat_history)
        answer = completion.choices[0].message.content
        role = completion.choices[0].message.role
        self.chat_history.append({"role": role, "content": answer})
        return answer, role

    def _generate_summary(self):
        """
        チャットの履歴から要約を生成し返す。

        要約する文字数は self._summary_length で指定する。
        ただし、GPTに文字数を指定して要約を生成させると、指定した文字数よりも多くなる場合がある。
        その場合、要約の文字数を self._summary_length に合わせ、最後に ... を追加する。
        :return: チャットの要約
        """

        # chat_history の先頭に要約の依頼を追加
        summary_request = {"role": "system",
                           "content": "あなたはユーザーの依頼を要約する役割を担います。"
                                      f"以下のユーザーの依頼を必ず全角{self._summary_length}文字以内で要約してください"}
        # GPTによる要約を取得
        messages = [summary_request, {"role": "user", "content": self._initial_prompt}]
        completion = openai.ChatCompletion.create(model=self.gpt_model, messages=messages)
        summary = completion.choices[0].message.content

        # 要約を調整
        if len(summary) > self._summary_length:
            return summary[:self._summary_length] + "..."
        else:
            return summary

    def chat_start(self):
        """
        AIアシスタントとユーザーとのチャットを開始し、チャットが終了したら要約を作成する。

        ユーザーからの入力を受け取り、AIアシスタントが応答を生成する。
        ユーザーが exit() と入力するとチャット終了。
        """

        self._give_role_to_gpt()
        self.gpt_model = self._choice_chat_model()
        if not self.gpt_model:
            print(f"{Fore.RED}エラーが発生しました。プログラムを終了します。{Fore.RESET}")
            exit()

        print(f"\nチャットを終了するには{Fore.GREEN} {self.EXIT_COMMAND} {Fore.RESET}と入力します。")
        user_prompt = self._input_user_prompt()

        # いきなり終了コマンドが入力された場合はプログラム終了
        if not self._initial_prompt and user_prompt == self.EXIT_COMMAND:
            exit()

        while user_prompt != self.EXIT_COMMAND:
            self.chat_history.append({"role": "user", "content": user_prompt})
            gpt_answer, gpt_role = self._fetch_gpt_answer()
            print(f"\n{Fore.GREEN}AIアシスタント:{Fore.RESET} {gpt_answer}")
            user_prompt = self._input_user_prompt()
        else:
            self.chat_summary = self._generate_summary()
