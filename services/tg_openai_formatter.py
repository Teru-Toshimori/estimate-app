import json
import logging
import os

from dotenv import load_dotenv
from openai import OpenAI


logger = logging.getLogger(__name__)


load_dotenv()



class TgOpenAIFormatter:
    """
    TG帳票 OpenAI整形処理

    OCR抽出後の値を
    Excel登録用JSONへ整形する

    """

    def __init__(self):

        api_key = os.getenv(
            "OPENAI_API_KEY"
        )


        if not api_key:

            raise RuntimeError(
                "OPENAI_API_KEY が設定されていません"
            )


        self.client = OpenAI(
            api_key=api_key
        )



    def format(
        self,
        data: dict
    ) -> dict:
        """
        抽出済みデータを整形

        Args:
            data:
                {
                    subject:"",
                    amount:"",
                    delivery_date:"",
                    department:""
                }

        Returns:
            JSON dict
        """


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

            response = self.client.responses.create(

                model="gpt-4.1-mini",

                input=[

                    {
                        "role":"user",

                        "content":[

                            {
                                "type":"input_text",

                                "text":prompt
                            }

                        ]
                    }

                ]

            )


            text = response.output_text


            logger.info(
                "OpenAI整形結果=%s",
                text
            )


            return json.loads(
                text
            )


        except Exception:

            logger.exception(
                "OpenAI整形失敗"
            )


            # 失敗時は元データ返却
            return data