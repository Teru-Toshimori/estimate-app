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
from services.ledger_reader import LedgerReader
from services.excel_writer import ExcelWriter
from models.estimate_data import EstimateData


class MainWindow(QMainWindow):

    def __init__(self):
        super().__init__()

        self.setWindowTitle("見積書作成ツール")
        self.resize(900, 700)

        self.setup_ui()

    def setup_ui(self):

        central = QWidget()
        self.setCentralWidget(central)

        main_layout = QVBoxLayout()
        central.setLayout(main_layout)

        # -----------------------------
        # PDF
        # -----------------------------
        pdf_label = QLabel("業務計画書(PDF)")

        self.pdf_edit = QLineEdit()

        self.pdf_button = QPushButton("参照")

        pdf_layout = QHBoxLayout()
        pdf_layout.addWidget(self.pdf_edit)
        pdf_layout.addWidget(self.pdf_button)

        # -----------------------------
        # Excel
        # -----------------------------
        excel_label = QLabel("管理台帳(Excel)")

        self.excel_edit = QLineEdit()

        self.excel_button = QPushButton("参照")

        excel_layout = QHBoxLayout()
        excel_layout.addWidget(self.excel_edit)
        excel_layout.addWidget(self.excel_button)

        # -----------------------------
        self.load_button = QPushButton("読み込み")

        # -----------------------------
        group = QGroupBox("抽出結果")

        form = QFormLayout()

        self.application_edit = QLineEdit()
        self.estimate_edit = QLineEdit()
        self.issue_edit = QLineEdit()
        self.department_edit = QLineEdit()
        self.subject_edit = QLineEdit()
        self.amount_edit = QLineEdit()
        self.delivery_edit = QLineEdit()
        self.outputs_edit = QTextEdit()

        form.addRow("申請書No", self.application_edit)
        form.addRow("見積番号", self.estimate_edit)
        form.addRow("発行日", self.issue_edit)
        form.addRow("依頼部署", self.department_edit)
        form.addRow("件名", self.subject_edit)
        form.addRow("委託金額", self.amount_edit)
        form.addRow("納期", self.delivery_edit)
        form.addRow("成果物名称", self.outputs_edit)

        group.setLayout(form)

        self.output_button = QPushButton("Excel出力")

        main_layout.addWidget(pdf_label)
        main_layout.addLayout(pdf_layout)

        main_layout.addWidget(excel_label)
        main_layout.addLayout(excel_layout)

        main_layout.addWidget(self.load_button)

        main_layout.addWidget(group)

        main_layout.addWidget(self.output_button)

        self.pdf_button.clicked.connect(self.select_pdf)
        self.excel_button.clicked.connect(self.select_excel)
        self.load_button.clicked.connect(self.load_pdf)

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
            "Excel (*.xlsx *.xls)"
        )

        if file_path:
            self.excel_edit.setText(file_path)

    def load_pdf(self):

        pdf_path = self.pdf_edit.text()
        excel_path = self.excel_edit.text()

        if not pdf_path:
            QMessageBox.warning(self, "確認", "業務計画書(PDF)を選択してください。")
            return

        if not excel_path:
            QMessageBox.warning(self, "確認", "管理台帳(Excel)を選択してください。")
            return

        try:
            reader = PDFReader()
            data = reader.parse(pdf_path)

            self.application_edit.setText(data.application_no)
            self.department_edit.setText(data.department)
            self.subject_edit.setText(data.subject)
            self.amount_edit.setText(data.amount)
            self.delivery_edit.setText(data.due_date)
            self.outputs_edit.setPlainText("\n".join(data.outputs))

            ledger = LedgerReader()

            ledger_data = ledger.find(
                excel_path,
                data.application_no
            )

            self.estimate_edit.setText(ledger_data.estimate_no)
            self.issue_edit.setText(ledger_data.issue_date)

            QMessageBox.information(self, "完了", "読み込みが完了しました。")

        except Exception as e:
            QMessageBox.critical(self, "エラー", f"読み込みに失敗しました。\n\n{e}")

    def export_excel(self):

        save_path, _ = QFileDialog.getSaveFileName(
            self,
            "見積書を保存",
            "",
            "Excelファイル (*.xlsx)"
        )

        if not save_path:
            return

        if not save_path.endswith(".xlsx"):
            save_path += ".xlsx"

        try:
            data = EstimateData(
                application_no=self.application_edit.text(),
                estimate_no=self.estimate_edit.text(),
                issue_date=self.issue_edit.text(),
                department=self.department_edit.text(),
                subject=self.subject_edit.text(),
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
                "Excel出力が完了しました。"
            )

        except Exception as e:
            QMessageBox.critical(
                self,
                "エラー",
                f"Excel出力に失敗しました。\n\n{e}"
            )