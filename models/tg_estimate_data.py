from dataclasses import dataclass, field


@dataclass
class TgEstimateData:
    """
    TG見積データモデル

    TG帳票から抽出した
    Excel出力用データを保持する
    """

    # 品名（件名）
    subject: str = ""

    # 金額
    amount: str = ""

    # 納期
    delivery_date: str = ""

    # 部署
    department: str = ""

    # 元PDFファイル名
    file_name: str = ""

    # OCR信頼度などの補足情報
    metadata: dict = field(
        default_factory=dict
    )