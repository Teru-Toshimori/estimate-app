import json
import logging
import os

from dotenv import load_dotenv
from openai import OpenAI


logger = logging.getLogger(__name__)


# .envファイル読込
load_dotenv()



class TgOpenAIFormatter:
    """
    TG帳票 OpenAI整形処理

    OCRで抽出した値を
    Excelへ登録しやすい形式へ整形する。

    処理の流れ

        OCR結果受取
            ↓
        OpenAIへ送信
            ↓
        JSON形式へ変換
            ↓
        呼び出し元へ返却
    """

    def __init__(self):

        # ---------------------------------
        # APIキー取得
        # ---------------------------------
        api_key = os.getenv(
            "OPENAI_API_KEY"
        )

        # APIキー未設定チェック
        if not api_key:

            raise RuntimeError(
                "OPENAI_API_KEY が設定されていません"
            )

        # OpenAIクライアント生成
        self.client = OpenAI(
            api_key=api_key
        )


    def format(
        self,
        data: dict
    ) -> dict:
        """
        OCR抽出済みデータを
        OpenAIで整形する

        Args
        ----
        data

            {
                subject:"",
                amount:"",
                delivery_date:"",
                department:""
            }

        Returns
        -------
        dict
        """

        # ---------------------------------
        # OpenAIへ渡すプロンプト作成
        # ---------------------------------
        prompt = f"""
あなたは帳票データ整形処理です。

以下のデータを指定形式に変換してください。

入力:
{json.dumps(
    data,
    ensure_ascii=False
)}

ルール:

・JSONのみ返す
・説明不要
・amountは数字のみ
・存在しない値は空文字

形式:

{{
 "subject":"",
 "amount":"",
 "delivery_date":"",
 "department":""
}}
"""

        try:

            # ---------------------------------
            # OpenAIへ送信
            # ---------------------------------
            response = self.client.responses.create(

                model="gpt-4.1-mini",

                input=[

                    {
                        "role": "user",

                        "content": [

                            {
                                "type": "input_text",

                                "text": prompt
                            }

                        ]
                    }

                ]

            )

            # ---------------------------------
            # OpenAI回答取得
            # ---------------------------------
            text = response.output_text

            logger.info(
                "OpenAI整形結果=%s",
                text
            )

            # ---------------------------------
            # JSONへ変換して返却
            # ---------------------------------
            return json.loads(
                text
            )

        except Exception:

            logger.exception(
                "OpenAI整形失敗"
            )

            # ---------------------------------
            # OpenAI失敗時は
            # 元データをそのまま返却
            # ---------------------------------
            return data