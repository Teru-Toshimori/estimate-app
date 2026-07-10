import os

from PySide6.QtWidgets import (
    QWidget,
    QMainWindow,
    QLabel,
    QPushButton,
    QLineEdit,
    QTextEdit,
    QVBoxLayout,
    QHBoxLayout,
    QFormLayout,
    QGroupBox,
    QFileDialog,
    QMessageBox,
)

from services.pdf_reader import PDFReader
from services.excel_writer import ExcelWriter
from services.ledger_writer import LedgerWriter
from models.estimate_data import EstimateData


class MainWindow(QMainWindow):

    def __init__(self):
        super().__init__()

        self.setWindowTitle("見積書作成ツール")
        self.resize(900, 750)

        self.setup_ui()

    def setup_ui(self):

        central = QWidget()
        self.setCentralWidget(central)

        main_layout = QVBoxLayout()
        central.setLayout(main_layout)

        # -----------------------------
        # PDF選択
        # -----------------------------
        pdf_label = QLabel("業務計画書(PDF)")

        self.pdf_edit = QLineEdit()
        self.pdf_button = QPushButton("参照")

        pdf_layout = QHBoxLayout()
        pdf_layout.addWidget(self.pdf_edit)
        pdf_layout.addWidget(self.pdf_button)

        # -----------------------------
        # 管理台帳選択
        # -----------------------------
        excel_label = QLabel("管理台帳(Excel)")

        self.excel_edit = QLineEdit()
        self.excel_button = QPushButton("参照")

        excel_layout = QHBoxLayout()
        excel_layout.addWidget(self.excel_edit)
        excel_layout.addWidget(self.excel_button)

        # -----------------------------
        # 操作ボタン
        # -----------------------------
        self.pdf_parse_button = QPushButton("PDF解析")
        self.ledger_write_button = QPushButton("台帳記入")
        self.output_button = QPushButton("Excel出力")

        # -----------------------------
        # 抽出結果
        # -----------------------------
        group = QGroupBox("抽出結果")

        form = QFormLayout()

        self.application_edit = QLineEdit()
        self.voucher_edit = QLineEdit()
        self.estimate_edit = QLineEdit()
        self.issue_edit = QLineEdit()
        self.department_edit = QLineEdit()
        self.subject_edit = QLineEdit()
        self.model_code_edit = QLineEdit()
        self.amount_edit = QLineEdit()
        self.delivery_edit = QLineEdit()
        self.outputs_edit = QTextEdit()

        form.addRow("申請書No", self.application_edit)
        form.addRow("伝票番号", self.voucher_edit)
        form.addRow("見積番号", self.estimate_edit)
        form.addRow("発行日", self.issue_edit)
        form.addRow("依頼部署", self.department_edit)
        form.addRow("件名", self.subject_edit)
        form.addRow("車種コード", self.model_code_edit)
        form.addRow("委託金額", self.amount_edit)
        form.addRow("納期", self.delivery_edit)
        form.addRow("成果物名称", self.outputs_edit)

        group.setLayout(form)

        # -----------------------------
        # レイアウト配置
        # -----------------------------
        main_layout.addWidget(pdf_label)
        main_layout.addLayout(pdf_layout)

        main_layout.addWidget(excel_label)
        main_layout.addLayout(excel_layout)

        main_layout.addWidget(self.pdf_parse_button)
        main_layout.addWidget(self.ledger_write_button)

        main_layout.addWidget(group)

        main_layout.addWidget(self.output_button)

        # -----------------------------
        # ボタンイベント
        # -----------------------------
        self.pdf_button.clicked.connect(self.select_pdf)
        self.excel_button.clicked.connect(self.select_excel)
        self.pdf_parse_button.clicked.connect(self.parse_pdf)
        self.ledger_write_button.clicked.connect(self.write_ledger)
        self.output_button.clicked.connect(self.export_excel)

    def select_pdf(self):

        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "業務計画書を選択",
            "",
            "PDF (*.pdf)"
        )

        if file_path:
            self.pdf_edit.setText(file_path)

    def select_excel(self):

        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "管理台帳を選択",
            "",
            "Excel (*.xlsx *.xlsm *.xls)"
        )

        if file_path:
            self.excel_edit.setText(file_path)

    def parse_pdf(self):

        pdf_path = self.pdf_edit.text()

        if not pdf_path:
            QMessageBox.warning(
                self,
                "確認",
                "業務計画書(PDF)を選択してください。"
            )
            return

        try:
            reader = PDFReader()
            data = reader.parse(pdf_path)

            voucher_no = os.path.splitext(
                os.path.basename(pdf_path)
            )[0]

            data.voucher_no = voucher_no

            self.application_edit.setText(data.application_no)
            self.voucher_edit.setText(data.voucher_no)
            self.department_edit.setText(data.department)
            self.subject_edit.setText(data.subject)
            self.model_code_edit.setText(data.model_code)
            self.amount_edit.setText(data.amount)
            self.delivery_edit.setText(data.due_date)
            self.outputs_edit.setPlainText("\n".join(data.outputs))

            QMessageBox.information(
                self,
                "完了",
                "PDF解析が完了しました。"
            )

        except Exception as e:
            QMessageBox.critical(
                self,
                "エラー",
                f"PDF解析に失敗しました。\n\n{e}"
            )

    def write_ledger(self):

        excel_path = self.excel_edit.text()

        if not excel_path:
            QMessageBox.warning(
                self,
                "確認",
                "管理台帳(Excel)を選択してください。"
            )
            return

        if not self.application_edit.text():
            QMessageBox.warning(
                self,
                "確認",
                "先にPDF解析を実行してください。"
            )
            return

        try:
            data = EstimateData(
                application_no=self.application_edit.text(),
                voucher_no=self.voucher_edit.text(),
                department=self.department_edit.text(),
                subject=self.subject_edit.text(),
                model_code=self.model_code_edit.text(),
                amount=self.amount_edit.text(),
                due_date=self.delivery_edit.text(),
                outputs=self.outputs_edit.toPlainText().splitlines(),
            )

            writer = LedgerWriter()

            result = writer.write(excel_path, data)

            self.estimate_edit.setText(result["estimate_no"])
            self.issue_edit.setText(result["issue_date"])

            QMessageBox.information(
                self,
                "完了",
                f"台帳記入が完了しました。\n\n"
                f"シート: {result['sheet_name']}\n"
                f"行: {result['row']}\n"
                f"見積番号: {result['estimate_no']}\n"
                f"発行日: {result['issue_date']}"
            )

        except Exception as e:
            QMessageBox.critical(
                self,
                "エラー",
                f"台帳記入に失敗しました。\n\n{e}"
            )

    def export_excel(self):

        save_path, _ = QFileDialog.getSaveFileName(
            self,
            "見積書を保存",
            "sample.xlsx",
            "Excelファイル (*.xlsx)"
        )

        if not save_path:
            return

        if not save_path.lower().endswith(".xlsx"):
            save_path += ".xlsx"

        try:
            data = EstimateData(
                application_no=self.application_edit.text(),
                voucher_no=self.voucher_edit.text(),
                estimate_no=self.estimate_edit.text(),
                issue_date=self.issue_edit.text(),
                department=self.department_edit.text(),
                subject=self.subject_edit.text(),
                model_code=self.model_code_edit.text(),
                amount=self.amount_edit.text(),
                due_date=self.delivery_edit.text(),
                outputs=self.outputs_edit.toPlainText().splitlines(),
            )

            writer = ExcelWriter()

            writer.write(
                "resources/TB_見積書フォーマット.xlsx",
                save_path,
                data
            )

            QMessageBox.information(
                self,
                "完了",
                "Excel出力とPDF出力が完了しました。"
            )

        except Exception as e:
            QMessageBox.critical(
                self,
                "エラー",
                f"Excel出力に失敗しました。\n\n{e}"
            )