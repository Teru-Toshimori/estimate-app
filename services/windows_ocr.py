import logging
import asyncio

from PIL import Image

from winrt.windows.graphics.imaging import (
    BitmapDecoder,
    BitmapPixelFormat,
    BitmapAlphaMode,
    SoftwareBitmap
)

from winrt.windows.media.ocr import OcrEngine


logger = logging.getLogger(__name__)


class WindowsOcr:
    """
    Windows OCR処理クラス

    処理:
        PIL Image
          ↓
        SoftwareBitmap
          ↓
        Windows OCR
          ↓
        OCR結果 + 座標情報
    """


    def __init__(self):

        self.ocr_engine = self._create_engine()



    def _create_engine(self):

        try:

            engine = (
                OcrEngine
                .try_create_from_user_profile_languages()
            )


            if engine is None:

                raise RuntimeError(
                    "OCRエンジンを作成できませんでした"
                )


            logger.info(
                "Windows OCR engine initialized"
            )


            return engine


        except Exception:

            logger.exception(
                "OCR engine initialization failed"
            )

            raise



    # =====================================
    # 通常OCR
    # =====================================
    def read(
        self,
        image: Image.Image
    ) -> str:
        """
        文字列のみ取得
        """

        result = asyncio.run(
            self._recognize_text(
                image
            )
        )

        return result



    # =====================================
    # 座標付きOCR
    # =====================================
    def read_with_position(
        self,
        image: Image.Image
    ):
        """
        OCR結果を座標付きで取得

        Returns:

        [
            {
                "text": "品名",
                "x":100,
                "y":200,
                "width":50,
                "height":20
            }
        ]

        """


        try:

            logger.info(
                "OCR position start"
            )


            result = asyncio.run(
                self._recognize_position(
                    image
                )
            )


            logger.info(
                "OCR item count: %s",
                len(result)
            )


            return result



        except Exception:

            logger.exception(
                "Position OCR failed"
            )

            raise




    async def _recognize_text(
        self,
        image
    ):

        bitmap = self._pil_to_software_bitmap(
            image
        )


        result = await (
            self.ocr_engine
            .recognize_async(bitmap)
        )


        logger.info(
            "OCR result:\n%s",
            result.text
        )


        return result.text




    async def _recognize_position(
        self,
        image
    ):

        bitmap = self._pil_to_software_bitmap(
            image
        )


        result = await (
            self.ocr_engine
            .recognize_async(bitmap)
        )


        items = []


        # 行単位
        for line in result.lines:


            for word in line.words:


                rect = word.bounding_rect


                items.append(
                    {
                        "text": word.text,

                        "x": rect.x,

                        "y": rect.y,

                        "width": rect.width,

                        "height": rect.height
                    }
                )


        return items




    def _pil_to_software_bitmap(
        self,
        image: Image.Image
    ):

        try:

            image = image.convert(
                "RGBA"
            )


            width, height = image.size


            pixels = image.load()


            # RGBA → BGRA
            bgra = bytearray()


            for y in range(height):

                for x in range(width):

                    r, g, b, a = pixels[x, y]


                    bgra.extend(
                        [
                            b,
                            g,
                            r,
                            a
                        ]
                    )



            bitmap = SoftwareBitmap(
                BitmapPixelFormat.BGRA8,
                width,
                height,
                BitmapAlphaMode.PREMULTIPLIED
            )


            bitmap.copy_from_buffer(
                bytes(bgra)
            )


            return bitmap



        except Exception:

            logger.exception(
                "PIL conversion failed"
            )

            raise