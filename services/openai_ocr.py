import base64
import io
import json
import logging
import os

from dotenv import load_dotenv
from openai import OpenAI
from PIL import Image

load_dotenv()

logger = logging.getLogger(__name__)


class OpenAIOcr:

    def __init__(self):

        api_key = os.getenv("OPENAI_API_KEY")

        if not api_key:
            raise RuntimeError("OPENAI_API_KEY が設定されていません。")

        self.client = OpenAI(api_key=api_key)

    def read(self, image: Image.Image) -> dict:

        buffer = io.BytesIO()

        image.save(buffer, format="PNG")

        image_base64 = base64.b64encode(
            buffer.getvalue()
        ).decode("utf-8")

        prompt = """
あなたは帳票OCRです。

画像から以下のみ抽出してください。

{
  "subject":"",
  "amount":"",
  "delivery_date":"",
  "department":""
}

ルール
・JSONのみ返す
・説明不要
・値が無ければ空文字
・amountは数字のみ
"""

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
                            "image_url": f"data:image/png;base64,{image_base64}",
                        },
                    ],
                }
            ],
        )

        text = response.output_text

        logger.info(text)

        return json.loads(text)