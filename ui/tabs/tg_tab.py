import os

from models.estimate_data import EstimateData
from PySide6.QtWidgets import (
    QFileDialog,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from services.tg_pdf_reader import TgPdfReader


class TgTab(QWidget):
    """
    TG向け見積書作成タブ
    """

    def __init__(self):
        super().__init__()

        # PDF解析結果を保持
        self.estimate_data: EstimateData | None = None

        self.setup_ui()

    def setup_ui(self):

        layout = QVBoxLayout(self)

        # タイトル
        title = QLabel("TG 見積書作成")
        title.setStyleSheet(
            "font-size:18px;font-weight:bold;"
        )
        layout.addWidget(title)

        # PDF
        pdf_label = QLabel("業務計画書（PDF）")

        self.pdf_edit = QLineEdit()
        self.pdf_edit.setPlaceholderText(
            "PDFを選択してください"
        )

        self.pdf_button = QPushButton("参照")

        pdf_layout = QHBoxLayout()
        pdf_layout.addWidget(self.pdf_edit)
        pdf_layout.addWidget(self.pdf_button)

        layout.addWidget(pdf_label)
        layout.addLayout(pdf_layout)

        # 管理台帳URL
        ledger_label = QLabel(
            "管理台帳URL（OneDrive／SharePoint）"
        )

        self.ledger_edit = QLineEdit()
        self.ledger_edit.setPlaceholderText(
            "https://..."
        )

        self.ledger_button = QPushButton("貼り付け")

        ledger_layout = QHBoxLayout()
        ledger_layout.addWidget(self.ledger_edit)
        ledger_layout.addWidget(self.ledger_button)

        layout.addWidget(ledger_label)
        layout.addLayout(ledger_layout)

        # 操作ボタン
        self.pdf_parse_button = QPushButton("PDF解析")
        self.ledger_write_button = QPushButton("台帳記入")

        button_layout = QHBoxLayout()
        button_layout.addWidget(self.pdf_parse_button)
        button_layout.addWidget(self.ledger_write_button)

        layout.addLayout(button_layout)

        # 出力
        self.output_button = QPushButton(
            "Excel・PDF出力"
        )

        layout.addWidget(self.output_button)

        # 状態
        self.status_label = QLabel("待機中")
        layout.addWidget(self.status_label)

        # イベント
        self.pdf_button.clicked.connect(
            self.select_pdf
        )

        self.ledger_button.clicked.connect(
            self.paste_url
        )

        self.pdf_parse_button.clicked.connect(
            self.pdf_parse
        )

        self.ledger_write_button.clicked.connect(
            self.write_ledger
        )

        self.output_button.clicked.connect(
            self.export_excel
        )

    # =====================================
    # PDF選択
    # =====================================
    def select_pdf(self):

        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "PDFを選択",
            "",
            "PDF (*.pdf)"
        )

        if file_path:
            self.pdf_edit.setText(file_path)

    # =====================================
    # URL貼り付け
    # =====================================
    def paste_url(self):

        QMessageBox.information(
            self,
            "未実装",
            "今後実装します。"
        )

    # =====================================
    # PDF解析
    # =====================================
    def pdf_parse(self):

        pdf_path = self.pdf_edit.text().strip()

        if not pdf_path:
            QMessageBox.warning(
                self,
                "確認",
                "PDFを選択してください。"
            )
            return

        reader = TgPdfReader()

        reader.parse(pdf_path)

        QMessageBox.information(
            self,
            "完了",
            "OCR結果をログへ出力しました。"
        )

    # =====================================
    # 台帳記入
    # =====================================
    def write_ledger(self):

        QMessageBox.information(
            self,
            "未実装",
            "TG用台帳記入を実装予定です。"
        )

    # =====================================
    # Excel・PDF出力
    # =====================================
    def export_excel(self):

        QMessageBox.information(
            self,
            "未実装",
            "TG用Excel出力を実装予定です。"
        )