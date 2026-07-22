import os
import logging
import base64
import json

from openai import OpenAI
from dotenv import load_dotenv


# .envファイル読込
load_dotenv()


logger = logging.getLogger(__name__)


class TgOpenAIVision:
    """
    TG仕様書 OpenAI Vision解析クラス
    
    処理の流れ
        PNG画像 → Base64変換 → OpenAI Vision送信
                → JSON解析 → dict返却
    """

    def __init__(self):
        """
        OpenAI クライアント初期化
        """
        
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


    # =====================================
    # 画像 → Base64
    # =====================================

    def encode_image(
        self,
        image_path: str
    ):
        """
        画像をBase64文字列へ変換する。

        OpenAI Vision APIへ画像を渡すため、
        PNGファイルをBase64へ変換する。
        """

        # ---------------------------------
        # ファイル存在確認
        # ---------------------------------
        if not os.path.exists(image_path):

            raise FileNotFoundError(
                image_path
            )

        # ---------------------------------
        # 画像読込
        # Base64へ変換
        # ---------------------------------
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
        self,
        content: str
    ):
        """
        OpenAIから返却されたJSONを解析する。

        Returns
        -------
        {
            "品名":"",
            "金額":""
        }
        """

        try:

            # ---------------------------------
            # JSON文字列→dict変換
            # ---------------------------------
            data = json.loads(
                content
            )

            # 必要項目だけ返却
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

            # JSON解析失敗時は空データ返却
            return {

                "品名": "",
                "金額": ""

            }



    # =====================================
    # OpenAI Vision
    # =====================================

    def extract_tg_data(
        self,
        image_path: str
    ):
        """
        TG仕様書画像から必要項目を抽出する。

        抽出項目

            ・品名（件名）
            ・予測金額

        処理の流れ

            PNG画像
                ↓
            Base64変換
                ↓
            OpenAI Vision送信
                ↓
            JSON取得
                ↓
            dict返却
        """

        logger.info(
            "OpenAI Vision解析開始"
        )

        try:

            # -----------------------------
            # 画像をBase64へ変換
            # -----------------------------
            image_base64 = self.encode_image(
                image_path
            )

            # -----------------------------
            # OpenAIへ送信するプロンプト
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
            # OpenAI Vision API実行
            # -----------------------------
            response = self.client.chat.completions.create(

                # Vision対応モデル
                model="gpt-4.1-mini",

                # JSON形式で返却
                response_format={

                    "type":
                    "json_object"

                },

                messages=[

                    {

                        "role":
                        "user",

                        "content":[

                            # 指示文
                            {

                                "type":
                                "text",

                                "text":
                                prompt

                            },

                            # 画像
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

                # 毎回同じ結果になるよう固定
                temperature=0

            )

            # ---------------------------------
            # OpenAI回答取得
            # ---------------------------------
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

            # ---------------------------------
            # JSON解析
            # ---------------------------------
            result = self.parse_json_response(
                content
            )

            logger.info(
                "抽出結果 : %s",
                result
            )

            # OCR結果返却
            return result

        except Exception:

            logger.exception(
                "OpenAI Vision解析失敗"
            )

            # エラー時は空データ返却
            return {

                "品名": "",
                "金額": ""

            }
