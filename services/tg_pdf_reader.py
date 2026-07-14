import logging
import re

from models.tg_estimate_data import TgEstimateData
from services.windows_ocr import WindowsOcr
from pypdfium2 import PdfDocument


logger = logging.getLogger(__name__)


class TgPdfReader:


    def parse(
        self,
        pdf_path: str
    ) -> TgEstimateData:


        data = TgEstimateData()

        pdf = PdfDocument(pdf_path)


        logger.info(
            "ページ数:%s",
            len(pdf)
        )


        ocr = WindowsOcr()


        all_items = []


        for page_no,page in enumerate(pdf):


            logger.info(
                "OCR開始 Page:%s",
                page_no + 1
            )


            bitmap = page.render(
                scale=3
            )


            image = bitmap.to_pil()


            items = ocr.read_with_position(
                image
            )


            logger.info(
                "OCR item count:%s",
                len(items)
            )


            all_items.extend(
                items
            )



        merged = self.merge_items(
            all_items
        )



        logger.info(
            "OCR解析開始"
        )


        data.subject = self.extract_subject(
            merged
        )


        data.amount = self.extract_amount(
            merged
        )



        logger.info(
            "品名 : %s",
            data.subject
        )


        logger.info(
            "予測金額 : %s",
            data.amount
        )


        return data



    def merge_items(
        self,
        items
    ):


        items = sorted(
            items,
            key=lambda x:
            (
                x["y"],
                x["x"]
            )
        )


        result=[]


        current=None


        for item in items:


            if current is None:

                current=item.copy()

                continue



            same_line = (
                abs(
                    item["y"]
                    -
                    current["y"]
                )
                <
                30
            )


            if same_line:


                current["text"] += item["text"]

                current["width"] += item["width"]


            else:

                result.append(
                    current
                )

                current=item.copy()



        if current:
            result.append(current)


        return result



    # =========================
    # 品名抽出
    # =========================

    def extract_subject(
        self,
        items
    ):


        # 品名ラベル探索

        label = None


        for item in items:

            text = self.normalize(
                item["text"]
            )


            if (
                "品名" in text
                or
                "件名" in text
            ):

                label = item

                break



        if label is None:

            logger.warning(
                "品名ラベルなし"
            )

            return None



        logger.info(
            "品名ラベル:%s x=%s y=%s",
            label["text"],
            label["x"],
            label["y"]
        )



        candidates=[]



        for item in items:


            # 同じ行

            same_y = abs(
                item["y"]
                -
                label["y"]
            ) < 40



            # ラベル右側

            right = (
                item["x"]
                >
                label["x"]
                +
                label["width"]
            )



            # 近すぎない

            distance = (
                item["x"]
                -
                (
                label["x"]
                +
                label["width"]
                )
            )


            if (
                same_y
                and
                right
                and
                distance < 500
            ):


                candidates.append(
                    item
                )



        candidates.sort(
            key=lambda x:x["x"]
        )


        value=""


        for c in candidates:

            value += c["text"]



        value=self.normalize(
            value
        )


        logger.info(
            "品名候補:%s",
            value
        )



        # 不要文字除去

        value=re.sub(
            r"(品名|件名|品番|S/W.*|希望納期.*)",
            "",
            value
        )


        return value.strip()



    # =========================
    # 金額抽出
    # =========================

    def extract_amount(
        self,
        items
    ):


        candidates=[]


        for item in items:


            text=item["text"]


            if "円" in text:

                candidates.append(
                    text
                )



        if not candidates:

            return None



        text="".join(
            candidates
        )


        logger.info(
            "金額候補:%s",
            text
        )


        m=re.search(
            r"\d[\d,]*",
            text
        )


        if m:


            return (
                m.group()
                .replace(",","")
            )



        return None



    def normalize(
        self,
        text
    ):


        replace={

            "ェイム":
            "エイム",

            "購人":
            "購入",

            "他開発品":
            "他開発品"

        }


        for k,v in replace.items():

            text=text.replace(
                k,
                v
            )


        return text