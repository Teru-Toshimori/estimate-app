from dataclasses import dataclass, field


@dataclass
class EstimateData:
    application_no: str = ""
    department: str = ""
    subject: str = ""
    outputs: list[str] = field(default_factory=list)
    amount: str = ""
    due_date: str = ""

    estimate_no: str = ""
    issue_date: str = ""