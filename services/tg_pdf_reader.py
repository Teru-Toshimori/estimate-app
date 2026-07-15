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
        画像化
          ↓
        OpenAI OCR
          ↓
        JSON
          ↓
        TgEstimateData
        """

        data = TgEstimateData()

        pdf = PdfDocument(pdf_path)

        logger.info(
            "ページ数: %s",
            len(pdf)
        )

        ocr = OpenAIOcr()

        for page_no, page in enumerate(pdf):

            logger.info(
                "OCR開始 Page:%s",
                page_no + 1
            )

            bitmap = page.render(
                scale=3
            )

            image = bitmap.to_pil()

            result = ocr.read(
                image
            )

            logger.info(
                "OCR結果: %s",
                result
            )

            # 最初に取得できた値を採用
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

            # 必要に応じて追加
            if hasattr(data, "delivery_date") and not getattr(data, "delivery_date"):
                data.delivery_date = result.get(
                    "delivery_date",
                    ""
                )

            if hasattr(data, "department") and not getattr(data, "department"):
                data.department = result.get(
                    "department",
                    ""
                )

        logger.info(
            "========== 抽出結果 =========="
        )

        logger.info(
            "品名 : %s",
            data.subject
        )

        logger.info(
            "予測金額 : %s",
            data.amount
        )

        if hasattr(data, "delivery_date"):
            logger.info(
                "納期 : %s",
                data.delivery_date
            )

        if hasattr(data, "department"):
            logger.info(
                "部署 : %s",
                data.department
            )

        return data