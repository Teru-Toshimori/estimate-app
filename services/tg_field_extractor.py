import logging
import re

from models.tg_estimate_data import TgEstimateData
from models.tg_ocr_result import TgOcrResult


logger = logging.getLogger(__name__)


class TgFieldExtractor:
    """
    TG仕様書 OCR結果から必要項目抽出

    抽出対象
    ・品名(件名)
    ・予測金額
    """


    def extract(
        self,
        ocr_results: list[TgOcrResult],
    ) -> TgEstimateData:

        data = TgEstimateData()


        data.subject = self._extract_right_value(
            ocr_results,
            [
                "品名(件名)",
                "品名（件名）",
            ]
        )


        data.amount = self._extract_money(
            ocr_results,
            [
                "予測金額",
            ]
        )


        logger.info(
            "========== 抽出結果 =========="
        )

        logger.info(
            "品名 : %s",
            data.subject
        )

        logger.info(
            "金額 : %s",
            data.amount
        )


        return data



    def _extract_right_value(
        self,
        ocr_results,
        labels
    ):
        """
        ラベル右側の文字取得
        """


        label = None


        # ラベル検索
        for item in ocr_results:

            text = (
                item.text
                .replace(" ", "")
                .replace("　", "")
            )


            for key in labels:

                if key in text:

                    label = item

                    logger.info(
                        "ラベル発見 : %s",
                        item.text
                    )

                    break


            if label:
                break



        if not label:

            logger.warning(
                "ラベルなし : %s",
                labels
            )

            return ""



        candidates = []


        for item in ocr_results:


            if item == label:
                continue


            # 右側
            if item.x <= label.x:
                continue


            # 同じ行
            if abs(item.y - label.y) > 30:
                continue


            candidates.append(item)



        if not candidates:

            logger.warning(
                "右側文字なし : %s",
                label.text
            )

            return ""



        # 左から近い順
        candidates.sort(
            key=lambda x:x.x
        )


        value = candidates[0].text.strip()


        logger.info(
            "%s -> %s",
            label.text,
            value
        )


        return value




    def _extract_money(
        self,
        ocr_results,
        labels
    ):
        """
        予測金額右側の金額取得
        """


        label = None


        for item in ocr_results:

            text = (
                item.text
                .replace(" ", "")
                .replace("　", "")
            )


            if any(
                key in text
                for key in labels
            ):

                label = item

                logger.info(
                    "金額ラベル発見 : %s",
                    item.text
                )

                break



        if not label:

            logger.warning(
                "金額ラベルなし"
            )

            return ""



        candidates=[]


        for item in ocr_results:


            if item.x <= label.x:
                continue


            if abs(item.y-label.y)>40:
                continue


            # 金額形式
            value = item.text.replace(
                ",",
                ""
            )


            if re.match(
                r"^\d+$",
                value
            ):

                candidates.append(item)



        if not candidates:

            logger.warning(
                "金額未検出"
            )

            return ""



        candidates.sort(
            key=lambda x:x.x
        )


        amount = candidates[0].text


        logger.info(
            "予測金額 -> %s",
            amount
        )


        return amount