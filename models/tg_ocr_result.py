from dataclasses import dataclass


@dataclass
class TgOcrResult:
    """
    TG帳票 OCR結果モデル

    PaddleOCRなどのOCR結果を保持する

    例:

    品名（件名）

    x=120
    y=200

    """

    # 認識した文字
    text: str


    # 左上X座標
    x: int


    # 左上Y座標
    y: int


    # 文字領域幅
    width: int


    # 文字領域高さ
    height: int


    # OCR信頼度
    confidence: float = 0.0



    def right(self) -> int:
        """
        文字領域の右端X座標

        例:
        
        x=100
        width=200

        → right=300

        """

        return self.x + self.width



    def bottom(self) -> int:
        """
        文字領域の下端Y座標
        """

        return self.y + self.height