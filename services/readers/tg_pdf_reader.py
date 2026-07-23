# services/tg_pdf_reader.py

import os
import logging

import fitz  # PyMuPDF

from services.ai.tg_openai_vision import extract_tg_data


logger = logging.getLogger(__name__)


class TgPdfReader:
    """
    TG仕様書PDF解析クラス

    処理の流れ

        PDF読込
            ↓
        1ページ目を画像化
            ↓
        OpenAI Visionへ送信
            ↓
        OCR結果取得
            ↓
        一時画像削除
    """

    def __init__(self):
        pass


    def convert_pdf_to_image(self, pdf_path):
        """
        PDFを画像化する

        TG仕様書は1ページ目のみを使用する。
        """

        logger.info("PDF画像変換開始")

        # ---------------------------------
        # PDFを開く
        # ---------------------------------
        doc = fitz.open(pdf_path)

        # ページ存在確認
        if len(doc) == 0:
            raise Exception(
                "PDFページが存在しません"
            )

        # ---------------------------------
        # 1ページ目取得
        # ---------------------------------
        page = doc[0]

        # ---------------------------------
        # 画像生成
        # dpi=200でOCRしやすい画質に変換
        # ---------------------------------
        pix = page.get_pixmap(
            dpi=200
        )

        # ---------------------------------
        # 一時画像ファイル名
        # 例)
        # sample.pdf
        #      ↓
        # sample_page1.png
        # ---------------------------------
        image_path = pdf_path.replace(
            ".pdf",
            "_page1.png"
        )

        # 画像保存
        pix.save(
            image_path
        )

        logger.info(
            "画像生成 : %s",
            image_path
        )

        return image_path


    def parse(self, pdf_path):
        """
        TG仕様書解析

        抽出項目
            ・品名（件名）
            ・予測金額

        戻り値

        {
            "品名": "...",
            "金額": "..."
        }
        """

        logger.info(
            "========== PDF解析開始 =========="
        )

        logger.info(
            "PDF : %s",
            pdf_path
        )

        # finallyで参照するため先に初期化
        image_path = None

        try:

            # -------------------------
            # PDFパスチェック
            # -------------------------

            if not pdf_path:

                raise ValueError(
                    "PDFファイルが指定されていません"
                )

            # -------------------------
            # ファイル存在確認
            # -------------------------

            if not os.path.exists(pdf_path):

                raise FileNotFoundError(
                    pdf_path
                )

            # -------------------------
            # PDF → PNG画像
            # -------------------------

            image_path = self.convert_pdf_to_image(
                pdf_path
            )

            # -------------------------
            # OpenAI Vision解析
            # -------------------------

            result = extract_tg_data(
                image_path
            )

            logger.info(
                "Vision結果 : %s",
                result
            )

            # -------------------------
            # OCR結果返却
            # -------------------------

            return result

        except Exception:

            logger.exception(
                "PDF解析失敗"
            )

            # エラー時は空データを返す
            return {
                "品名": "",
                "金額": ""
            }

        finally:

            # ---------------------------------
            # 一時画像削除
            #
            # OpenAI Vision送信用に作成したPNGは
            # 不要になるため削除する
            # ---------------------------------
            if image_path and os.path.exists(image_path):

                try:

                    os.remove(image_path)

                    logger.info(
                        "一時画像削除 : %s",
                        image_path
                    )

                except Exception:

                    logger.exception(
                        "一時画像削除失敗"
                    )

            logger.info("========== PDF解析終了 ==========")