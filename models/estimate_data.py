from dataclasses import dataclass, field


@dataclass
class EstimateData:

    # =============================
    # PDFから取得する情報
    # =============================
    application_no: str = ""
    voucher_no: str = ""
    department: str = ""
    subject: str = ""
    model_code: str = ""

    outputs: list[str] = field(default_factory=list)

    amount: str = ""
    due_date: str = ""

    # =============================
    # 台帳記入・Excel出力で使用
    # =============================
    estimate_no: str = ""
    issue_date: str = ""