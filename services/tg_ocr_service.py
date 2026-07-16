import logging
import pytesseract

from PIL import Image


logger = logging.getLogger(__name__)


# Tesseract設定
TESSERACT_PATH = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

TESSDATA_PATH = r"C:\Program Files\Tesseract-OCR\tessdata"


pytesseract.pytesseract.tesseract_cmd = (
    TESSERACT_PATH
)


logger.info(
    f"Tesseract : {TESSERACT_PATH}"
)

logger.info(
    f"tessdata : {TESSDATA_PATH}"
)



def extract_ocr_coordinates(image_path):
    """
    OCR文字 + 座標取得

    戻り値:

    [
      {
        "text":"品名(件名)",
        "left":100,
        "top":200,
        "width":80,
        "height":20
      }
    ]

    """


    logger.info(
        "========== OCR開始 =========="
    )

    logger.info(
        f"Image : {image_path}"
    )


    try:

        image = Image.open(
            image_path
        )


        # 座標付きOCR
        data = pytesseract.image_to_data(
            image,
            lang="jpn",
            output_type=pytesseract.Output.DICT
        )


        results = []


        count = len(data["text"])


        logger.info(
            f"OCR検出数 : {count}"
        )


        for i in range(count):

            text = data["text"][i].strip()


            # 空文字除外
            if not text:
                continue


            item = {

                "text": text,

                "left": int(
                    data["left"][i]
                ),

                "top": int(
                    data["top"][i]
                ),

                "width": int(
                    data["width"][i]
                ),

                "height": int(
                    data["height"][i]
                ),

                "confidence": float(
                    data["conf"][i]
                )
            }


            results.append(item)



        logger.info(
            f"OCR取得数 : {len(results)}"
        )


        # 確認ログ
        logger.info(
            "========== OCR座標一覧 =========="
        )


        for r in results:

            logger.info(
                f"{r['text']} "
                f"x={r['left']} "
                f"y={r['top']}"
            )



        return results



    except Exception:

        logger.exception(
            "OCR解析失敗"
        )

        return []