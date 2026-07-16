import os
import logging
import base64
import json

from openai import OpenAI
from dotenv import load_dotenv


load_dotenv()


logger = logging.getLogger(__name__)


# =====================================
# OpenAI設定
# =====================================

OPENAI_API_KEY = os.getenv(
    "OPENAI_API_KEY"
)


if not OPENAI_API_KEY:

    raise RuntimeError(
        "OPENAI_API_KEY が設定されていません"
    )


client = OpenAI(
    api_key=OPENAI_API_KEY
)



# =====================================
# 画像 → Base64
# =====================================

def encode_image(
    image_path: str
):
    """
    画像をBase64へ変換
    """

    if not os.path.exists(image_path):

        raise FileNotFoundError(
            image_path
        )


    with open(
        image_path,
        "rb"
    ) as f:

        return base64.b64encode(
            f.read()
        ).decode(
            "utf-8"
        )



# =====================================
# JSON整形
# =====================================

def parse_json_response(
    content: str
):
    """
    OpenAI回答JSON解析
    """

    try:

        data = json.loads(
            content
        )


        return {

            "品名":
                data.get(
                    "品名",
                    ""
                ),

            "金額":
                data.get(
                    "金額",
                    ""
                )
        }


    except Exception:


        logger.exception(
            "JSON解析失敗 : %s",
            content
        )


        return {

            "品名":"",
            "金額":""

        }



# =====================================
# OpenAI Vision
# =====================================

def extract_tg_data(
    image_path: str
):
    """
    TG仕様書画像から必要項目抽出

    抽出:
        品名（件名）
        予測金額

    """

    logger.info(
        "OpenAI Vision解析開始"
    )


    try:


        # -----------------------------
        # 画像取得
        # -----------------------------

        image_base64 = encode_image(
            image_path
        )


        # -----------------------------
        # Prompt
        # -----------------------------

        prompt = """

帳票画像から以下2項目だけ抽出してください。

【抽出項目】

1.
品名（件名）

取得位置:
「品名（件名）」
または
「品名(件名)」
の右側


例:
品名（件名） 設計・製図委託


2.
予測金額

取得位置:
「予測金額」
の右側にある金額


必ずJSON形式で返してください。

{
 "品名":"",
 "金額":""
}

存在しない場合は空文字。

"""


        # -----------------------------
        # OpenAI API
        # -----------------------------

        response = client.chat.completions.create(

            model="gpt-4.1-mini",


            response_format={

                "type":
                "json_object"

            },


            messages=[

                {

                    "role":
                    "user",


                    "content":[


                        {

                            "type":
                            "text",

                            "text":
                            prompt

                        },


                        {

                            "type":
                            "image_url",


                            "image_url":{

                                "url":
                                f"data:image/png;base64,{image_base64}"

                            }

                        }

                    ]

                }

            ],


            temperature=0

        )



        content = (
            response
            .choices[0]
            .message
            .content
        )


        logger.info(
            "OpenAI返却 : %s",
            content
        )



        # -----------------------------
        # JSON変換
        # -----------------------------

        result = parse_json_response(
            content
        )


        logger.info(
            "抽出結果 : %s",
            result
        )


        return result



    except Exception:


        logger.exception(
            "OpenAI Vision解析失敗"
        )


        return {

            "品名":"",
            "金額":""

        }