import logging

from pypdfium2 import PdfDocument

from models.tg_estimate_data import TgEstimateData
from services.openai_ocr import OpenAIOcr


logger = logging.getLogger(__name__)


class TgPdfReader:

    def parse(
        self,
        pdf_path: str
    ) -> TgEstimateData:
        """
        PDF解析

        PDF
          ↓
        pypdfium2
          ↓
        高解像度画像化
          ↓
        OCR対象領域切り抜き
          ↓
        OpenAI OCR
          ↓
        JSON
          ↓
        TgEstimateData
        """

        data = TgEstimateData()


        logger.info(
            "PDF解析開始 : %s",
            pdf_path
        )


        try:

            pdf = PdfDocument(
                pdf_path
            )


            logger.info(
                "ページ数 : %s",
                len(pdf)
            )


            ocr = OpenAIOcr()


            for page_no, page in enumerate(pdf):

                logger.info(
                    "OCR開始 Page:%s",
                    page_no + 1
                )


                try:

                    # ---------------------------------
                    # PDF → 高解像度画像
                    # scale=5 約360dpi相当
                    # ---------------------------------
                    bitmap = page.render(
                        scale=5
                    )


                    image = bitmap.to_pil()


                    logger.info(
                        "画像サイズ : %s",
                        image.size
                    )


                    # ---------------------------------
                    # OCR対象範囲
                    #
                    # 現在:
                    # ページ上部40%
                    #
                    # 帳票レイアウト確認後、
                    # 座標固定推奨
                    # ---------------------------------
                    width, height = image.size


                    ocr_image = image.crop(
                        (
                            0,
                            0,
                            width,
                            int(height * 0.4)
                        )
                    )


                    logger.info(
                        "OCR画像サイズ : %s",
                        ocr_image.size
                    )


                    result = ocr.read(
                        ocr_image
                    )


                    logger.info(
                        "OCR結果 Page:%s %s",
                        page_no + 1,
                        result
                    )


                    # ---------------------------------
                    # 最初に取得できた値を採用
                    # ---------------------------------

                    if not data.subject:

                        data.subject = result.get(
                            "subject",
                            ""
                        )


                    if not data.amount:

                        data.amount = result.get(
                            "amount",
                            ""
                        )


                    if hasattr(
                        data,
                        "delivery_date"
                    ):

                        if not getattr(
                            data,
                            "delivery_date"
                        ):

                            data.delivery_date = result.get(
                                "delivery_date",
                                ""
                            )


                    if hasattr(
                        data,
                        "department"
                    ):

                        if not getattr(
                            data,
                            "department"
                        ):

                            data.department = result.get(
                                "department",
                                ""
                            )


                except Exception:

                    logger.exception(
                        "Page:%s OCR処理失敗",
                        page_no + 1
                    )

                    continue


            pdf.close()


        except Exception:

            logger.exception(
                "PDF解析失敗"
            )

            return data



        # ---------------------------------
        # 最終結果ログ
        # ---------------------------------

        logger.info(
            "========== 抽出結果 =========="
        )


        logger.info(
            "件名 : %s",
            data.subject
        )


        logger.info(
            "金額 : %s",
            data.amount
        )


        if hasattr(
            data,
            "delivery_date"
        ):

            logger.info(
                "納期 : %s",
                data.delivery_date
            )


        if hasattr(
            data,
            "department"
        ):

            logger.info(
                "部署 : %s",
                data.department
            )


        logger.info(
            "PDF解析終了"
        )


        return data