import logging

import cv2
import numpy as np

from PIL import Image


logger = logging.getLogger(__name__)


class TgImageProcessor:
    """
    TG帳票 画像前処理

    目的:
        OCR精度向上

    処理:
        ・グレースケール化
        ・ノイズ除去
        ・二値化
        ・コントラスト補正
    """

    def process(
        self,
        image: Image.Image
    ) -> Image.Image:
        """
        OCR用画像へ変換

        Args:
            image:
                PIL Image

        Returns:
            PIL Image
        """

        try:

            logger.info(
                "TG画像補正開始 size=%s",
                image.size
            )


            # PIL → OpenCV
            img = np.array(
                image
            )


            # RGB → BGR
            img = cv2.cvtColor(
                img,
                cv2.COLOR_RGB2BGR
            )


            # -------------------------
            # グレースケール
            # -------------------------
            gray = cv2.cvtColor(
                img,
                cv2.COLOR_BGR2GRAY
            )


            # -------------------------
            # ノイズ除去
            # -------------------------
            denoise = cv2.medianBlur(
                gray,
                3
            )


            # -------------------------
            # コントラスト補正
            # -------------------------
            clahe = cv2.createCLAHE(
                clipLimit=2.0,
                tileGridSize=(8, 8)
            )

            enhanced = clahe.apply(
                denoise
            )


            # -------------------------
            # 二値化
            # -------------------------
            binary = cv2.adaptiveThreshold(
                enhanced,
                255,
                cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
                cv2.THRESH_BINARY,
                31,
                15
            )


            # OpenCV → PIL
            result = Image.fromarray(
                binary
            )


            logger.info(
                "TG画像補正完了"
            )


            return result



        except Exception:

            logger.exception(
                "TG画像補正失敗"
            )

            # 失敗時は元画像を返す
            return image