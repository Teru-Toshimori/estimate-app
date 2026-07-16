# services/tg_pdf_reader.py

import os
import logging

import fitz  # PyMuPDF

from services.tg_openai_vision import extract_tg_data


logger = logging.getLogger(__name__)


class TgPdfReader:

    def __init__(self):
        pass


    def convert_pdf_to_image(self, pdf_path):
        """
        PDFを画像化する
        """

        logger.info("PDF画像変換開始")

        doc = fitz.open(pdf_path)

        if len(doc) == 0:
            raise Exception(
                "PDFページが存在しません"
            )

        page = doc[0]

        pix = page.get_pixmap(
            dpi=200
        )


        image_path = pdf_path.replace(
            ".pdf",
            "_page1.png"
        )


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

        抽出項目:
        ・品名（件名）
        ・予測金額

        """

        logger.info(
            "========== PDF解析開始 =========="
        )

        logger.info(
            "PDF : %s",
            pdf_path
        )


        try:

            # -------------------------
            # PDFチェック
            # -------------------------

            if not pdf_path:

                raise ValueError(
                    "PDFファイルが指定されていません"
                )


            if not os.path.exists(pdf_path):

                raise FileNotFoundError(
                    pdf_path
                )


            # -------------------------
            # PDF → 画像
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
            # 戻り値
            # -------------------------

            return result



        except Exception:

            logger.exception(
                "PDF解析失敗"
            )


            return {
                "品名": "",
                "金額": ""
            }



        finally:

            logger.info(
                "========== PDF解析終了 =========="
            )