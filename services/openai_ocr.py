import base64
import io
import json
import logging
import os
import re

from dotenv import load_dotenv
from openai import OpenAI
from PIL import Image


load_dotenv()

logger = logging.getLogger(__name__)


class OpenAIOcr:

    def __init__(self):

        api_key = os.getenv("OPENAI_API_KEY")

        if not api_key:
            raise RuntimeError(
                "OPENAI_API_KEY が設定されていません。"
            )

        self.client = OpenAI(
            api_key=api_key
        )


    def read(self, image: Image.Image) -> dict:
        """
        画像から帳票項目をOCR抽出する

        Args:
            image:
                PIL Image

        Returns:
            dict:
                OCR結果
        """

        try:

            logger.info(
                "OCR開始 image_size=%s",
                image.size
            )


            # -----------------------------
            # Image → Base64
            # -----------------------------
            buffer = io.BytesIO()

            image.save(
                buffer,
                format="PNG"
            )

            image_base64 = base64.b64encode(
                buffer.getvalue()
            ).decode("utf-8")


            # -----------------------------
            # OCR用プロンプト
            # -----------------------------
            prompt = """
あなたは帳票文字認識専用OCRエンジンです。

画像内に存在する文字をそのまま読み取ってください。

重要ルール:
・推測は禁止
・意味による補完は禁止
・似た意味の文字へ変更禁止
・別項目の値を使用禁止
・画像内に存在しない情報は空文字
・読めない場合は空文字

以下JSON形式のみ返してください。

{
  "subject":"",
  "amount":"",
  "delivery_date":"",
  "department":""
}

追加ルール:
・JSON以外の文章は禁止
・説明は禁止
・amountは数字のみ
"""


            # -----------------------------
            # OpenAI Vision OCR
            # -----------------------------
            response = self.client.responses.create(

                model="gpt-4.1",

                input=[
                    {
                        "role": "user",

                        "content": [

                            {
                                "type": "input_text",
                                "text": prompt,
                            },

                            {
                                "type": "input_image",

                                "image_url":
                                    f"data:image/png;base64,{image_base64}",
                            },
                        ],
                    }
                ],
            )


            result_text = response.output_text.strip()


            logger.info(
                "OpenAI response=%s",
                result_text
            )


            # -----------------------------
            # JSON部分のみ取得
            # -----------------------------
            json_text = self._extract_json(
                result_text
            )


            result = json.loads(
                json_text
            )


            logger.info(
                "OCR結果=%s",
                json.dumps(
                    result,
                    ensure_ascii=False
                )
            )


            return result



        except json.JSONDecodeError:

            logger.exception(
                "OCR結果JSON解析失敗"
            )

            return self._empty_result()



        except Exception:

            logger.exception(
                "OCR処理中にエラー発生"
            )

            return self._empty_result()



    def _extract_json(
        self,
        text: str
    ) -> str:
        """
        OpenAI回答からJSON部分だけ抽出
        """

        match = re.search(
            r"\{.*\}",
            text,
            re.DOTALL
        )

        if not match:

            raise ValueError(
                "JSONが見つかりません"
            )


        return match.group()



    def _empty_result(self) -> dict:
        """
        OCR失敗時の戻り値
        """

        return {

            "subject": "",

            "amount": "",

            "delivery_date": "",

            "department": ""

        }