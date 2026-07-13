import os

from PySide6.QtCore import QThread, Signal
from PySide6.QtWidgets import (
    QApplication,
    QFileDialog,
    QFormLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from models.estimate_data import EstimateData
from services.excel_writer import ExcelWriter
from services.pdf_reader import PDFReader
from ui.device_login_dialog import DeviceLoginDialog
from workers.ledger_update_worker import LedgerUpdateWorker


class TokuchoTbTab(QWidget):
    """
    特調TB向けの見積書作成・台帳記入機能。
    """

    device_login_requested = Signal(dict)
    status_changed = Signal(str)

    def __init__(self, parent=None):
        super().__init__(parent)

        self.ledger_thread = None
        self.ledger_worker = None

        self.setup_ui()

        self.device_login_requested.connect(
            self.show_device_login
        )

    def setup_ui(self):

        main_layout = QVBoxLayout(self)

        # =====================================
        # タイトル
        # =====================================
        title_label = QLabel("特調TB 見積書作成")
        title_label.setStyleSheet(
            "font-size: 18px; font-weight: bold;"
        )

        main_layout.addWidget(title_label)

        # =====================================
        # 業務計画書PDF
        # =====================================
        pdf_label = QLabel("業務計画書（PDF）")

        self.pdf_edit = QLineEdit()
        self.pdf_edit.setPlaceholderText(
            "業務計画書PDFを選択してください"
        )

        self.pdf_button = QPushButton("参照")

        pdf_layout = QHBoxLayout()
        pdf_layout.addWidget(self.pdf_edit)
        pdf_layout.addWidget(self.pdf_button)

        main_layout.addWidget(pdf_label)
        main_layout.addLayout(pdf_layout)

        # =====================================
        # OneDrive／SharePoint管理台帳URL
        # =====================================
        ledger_url_label = QLabel(
            "管理台帳URL（OneDrive／SharePoint）"
        )

        self.excel_edit = QLineEdit()
        self.excel_edit.setPlaceholderText(
            "https://会社名.sharepoint.com/..."
        )

        self.excel_button = QPushButton("貼り付け")

        ledger_url_layout = QHBoxLayout()
        ledger_url_layout.addWidget(self.excel_edit)
        ledger_url_layout.addWidget(self.excel_button)

        main_layout.addWidget(ledger_url_label)
        main_layout.addLayout(ledger_url_layout)

        # =====================================
        # 操作ボタン
        # =====================================
        self.pdf_parse_button = QPushButton("PDF解析")
        self.ledger_write_button = QPushButton("台帳記入")

        operation_layout = QHBoxLayout()
        operation_layout.addWidget(self.pdf_parse_button)
        operation_layout.addWidget(self.ledger_write_button)

        main_layout.addLayout(operation_layout)

        # =====================================
        # 抽出結果
        # =====================================
        result_group = QGroupBox("抽出結果")
        result_form = QFormLayout()

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

        self.outputs_edit.setMinimumHeight(130)

        result_form.addRow(
            "申請書No",
            self.application_edit,
        )

        result_form.addRow(
            "伝票番号",
            self.voucher_edit,
        )

        result_form.addRow(
            "見積番号",
            self.estimate_edit,
        )

        result_form.addRow(
            "発行日",
            self.issue_edit,
        )

        result_form.addRow(
            "依頼部署",
            self.department_edit,
        )

        result_form.addRow(
            "件名",
            self.subject_edit,
        )

        result_form.addRow(
            "車種コード",
            self.model_code_edit,
        )

        result_form.addRow(
            "委託金額",
            self.amount_edit,
        )

        result_form.addRow(
            "納期",
            self.delivery_edit,
        )

        result_form.addRow(
            "成果物名称",
            self.outputs_edit,
        )

        result_group.setLayout(result_form)

        main_layout.addWidget(result_group)

        # =====================================
        # Excel・PDF出力
        # =====================================
        self.output_button = QPushButton(
            "Excel・PDF出力"
        )

        main_layout.addWidget(self.output_button)

        # =====================================
        # 状態表示
        # =====================================
        self.status_label = QLabel("待機中")
        self.status_label.setStyleSheet(
            "color: #555555;"
        )

        main_layout.addWidget(self.status_label)

        # =====================================
        # イベント接続
        # =====================================
        self.pdf_button.clicked.connect(
            self.select_pdf
        )

        self.excel_button.clicked.connect(
            self.paste_excel_url
        )

        self.pdf_parse_button.clicked.connect(
            self.parse_pdf
        )

        self.ledger_write_button.clicked.connect(
            self.write_ledger
        )

        self.output_button.clicked.connect(
            self.export_excel
        )

        self.status_changed.connect(
            self.set_status
        )

    # =====================================
    # 状態表示
    # =====================================
    def set_status(self, message: str):

        self.status_label.setText(message)

    # =====================================
    # PDF選択
    # =====================================
    def select_pdf(self):

        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "業務計画書を選択",
            "",
            "PDFファイル (*.pdf)"
        )

        if file_path:
            self.pdf_edit.setText(file_path)

    # =====================================
    # URL貼り付け
    # =====================================
    def paste_excel_url(self):

        clipboard = QApplication.clipboard()
        url = clipboard.text().strip()

        if not url:
            QMessageBox.warning(
                self,
                "確認",
                "クリップボードにURLがありません。"
            )
            return

        if not url.startswith(
            ("https://", "http://")
        ):
            QMessageBox.warning(
                self,
                "確認",
                "クリップボードの内容がURLではありません。"
            )
            return

        self.excel_edit.setText(url)

    # =====================================
    # PDF解析
    # =====================================
    def parse_pdf(self):

        pdf_path = self.pdf_edit.text().strip()

        if not pdf_path:
            QMessageBox.warning(
                self,
                "確認",
                "業務計画書（PDF）を選択してください。"
            )
            return

        if not os.path.exists(pdf_path):
            QMessageBox.warning(
                self,
                "確認",
                "指定されたPDFファイルが見つかりません。\n\n"
                f"{pdf_path}"
            )
            return

        try:
            self.set_status(
                "PDFを解析しています..."
            )

            reader = PDFReader()
            data = reader.parse(pdf_path)

            voucher_no = os.path.splitext(
                os.path.basename(pdf_path)
            )[0]

            data.voucher_no = voucher_no

            self.application_edit.setText(
                data.application_no
            )

            self.voucher_edit.setText(
                data.voucher_no
            )

            self.department_edit.setText(
                data.department
            )

            self.subject_edit.setText(
                data.subject
            )

            self.model_code_edit.setText(
                data.model_code
            )

            self.amount_edit.setText(
                data.amount
            )

            self.delivery_edit.setText(
                data.due_date
            )

            self.outputs_edit.setPlainText(
                "\n".join(data.outputs)
            )

            self.estimate_edit.clear()
            self.issue_edit.clear()

            self.set_status(
                "PDF解析が完了しました。"
            )

            QMessageBox.information(
                self,
                "完了",
                "PDF解析が完了しました。"
            )

        except Exception as error:
            self.set_status(
                "PDF解析に失敗しました。"
            )

            QMessageBox.critical(
                self,
                "エラー",
                "PDF解析に失敗しました。\n\n"
                f"{error}"
            )

    # =====================================
    # 管理台帳記入
    # =====================================
    def write_ledger(self):

        share_url = self.excel_edit.text().strip()

        if not share_url:
            QMessageBox.warning(
                self,
                "確認",
                "管理台帳URLを入力してください。"
            )
            return

        if not share_url.startswith(
            ("https://", "http://")
        ):
            QMessageBox.warning(
                self,
                "確認",
                "管理台帳URLの形式が正しくありません。"
            )
            return

        if not self.application_edit.text().strip():
            QMessageBox.warning(
                self,
                "確認",
                "先にPDF解析を実行してください。"
            )
            return

        if not self.department_edit.text().strip():
            QMessageBox.warning(
                self,
                "確認",
                "依頼部署が入力されていません。"
            )
            return

        if (
            self.ledger_thread is not None
            and self.ledger_thread.isRunning()
        ):
            QMessageBox.warning(
                self,
                "確認",
                "現在、台帳記入処理を実行中です。"
            )
            return

        confirm = QMessageBox.question(
            self,
            "台帳記入の確認",
            "OneDrive／SharePoint上の管理台帳へ"
            "新しい行を追加します。\n\n"
            "実行してよろしいですか？",
            (
                QMessageBox.StandardButton.Yes
                | QMessageBox.StandardButton.No
            ),
            QMessageBox.StandardButton.No,
        )

        if confirm != QMessageBox.StandardButton.Yes:
            return

        try:
            data = self.create_estimate_data()

            self.start_ledger_update(
                share_url=share_url,
                data=data,
            )

        except Exception as error:
            QMessageBox.critical(
                self,
                "エラー",
                "台帳記入処理を開始できませんでした。\n\n"
                f"{error}"
            )

    # =====================================
    # バックグラウンド更新開始
    # =====================================
    def start_ledger_update(
        self,
        share_url: str,
        data: EstimateData,
    ):

        self.ledger_write_button.setEnabled(False)
        self.ledger_write_button.setText(
            "台帳記入中..."
        )

        self.set_status(
            "管理台帳の更新を開始しています..."
        )

        self.ledger_thread = QThread(self)

        self.ledger_worker = LedgerUpdateWorker(
            share_url=share_url,
            data=data,
            device_flow_callback=(
                self.request_device_login
            ),
        )

        self.ledger_worker.moveToThread(
            self.ledger_thread
        )

        self.ledger_thread.started.connect(
            self.ledger_worker.run
        )

        self.ledger_worker.progress.connect(
            self.set_status
        )

        self.ledger_worker.finished.connect(
            self.on_ledger_update_finished
        )

        self.ledger_worker.failed.connect(
            self.on_ledger_update_failed
        )

        self.ledger_worker.finished.connect(
            self.ledger_thread.quit
        )

        self.ledger_worker.failed.connect(
            self.ledger_thread.quit
        )

        self.ledger_thread.finished.connect(
            self.ledger_worker.deleteLater
        )

        self.ledger_thread.finished.connect(
            self.clear_ledger_thread
        )

        self.ledger_thread.start()

    # =====================================
    # 認証要求
    # =====================================
    def request_device_login(self, flow: dict):

        self.device_login_requested.emit(flow)

    # =====================================
    # 認証ダイアログ
    # =====================================
    def show_device_login(self, flow: dict):

        dialog = DeviceLoginDialog(
            flow=flow,
            parent=self,
        )

        dialog.exec()

    # =====================================
    # 更新成功
    # =====================================
    def on_ledger_update_finished(
        self,
        result: dict,
    ):

        self.estimate_edit.setText(
            result["estimate_no"]
        )

        self.issue_edit.setText(
            result["issue_date"]
        )

        self.set_status(
            "管理台帳の更新が完了しました。"
        )

        QMessageBox.information(
            self,
            "完了",
            "OneDrive／SharePoint上の管理台帳へ"
            "記入しました。\n\n"
            f"ファイル："
            f"{result.get('remote_name', '')}\n"
            f"シート：{result['sheet_name']}\n"
            f"行：{result['row']}\n"
            f"見積番号：{result['estimate_no']}\n"
            f"発行日：{result['issue_date']}\n"
            f"発行者：{result.get('issuer_name', '')}"
        )

    # =====================================
    # 更新失敗
    # =====================================
    def on_ledger_update_failed(
        self,
        error_message: str,
    ):

        self.set_status(
            "管理台帳の更新に失敗しました。"
        )

        QMessageBox.critical(
            self,
            "エラー",
            "OneDrive／SharePointの台帳記入に"
            "失敗しました。\n\n"
            f"{error_message}"
        )

    # =====================================
    # スレッド終了
    # =====================================
    def clear_ledger_thread(self):

        self.ledger_write_button.setEnabled(True)
        self.ledger_write_button.setText(
            "台帳記入"
        )

        thread = self.ledger_thread

        self.ledger_worker = None
        self.ledger_thread = None

        if thread is not None:
            thread.deleteLater()

    # =====================================
    # Excel・PDF出力
    # =====================================
    def export_excel(self):

        if not self.application_edit.text().strip():
            QMessageBox.warning(
                self,
                "確認",
                "先にPDF解析を実行してください。"
            )
            return

        if not self.estimate_edit.text().strip():
            QMessageBox.warning(
                self,
                "確認",
                "先に台帳記入を実行してください。"
            )
            return

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

        self.output_button.setEnabled(False)
        self.output_button.setText(
            "出力中..."
        )

        QApplication.processEvents()

        try:
            data = self.create_estimate_data()

            writer = ExcelWriter()

            writer.write(
                "resources/TB_見積書フォーマット.xlsx",
                save_path,
                data
            )

            pdf_path = os.path.splitext(
                save_path
            )[0] + ".pdf"

            self.set_status(
                "Excel・PDF出力が完了しました。"
            )

            QMessageBox.information(
                self,
                "完了",
                "Excel出力とPDF出力が完了しました。\n\n"
                f"Excel：{save_path}\n"
                f"PDF：{pdf_path}"
            )

        except Exception as error:
            self.set_status(
                "Excel・PDF出力に失敗しました。"
            )

            QMessageBox.critical(
                self,
                "エラー",
                "Excel・PDF出力に失敗しました。\n\n"
                f"{error}"
            )

        finally:
            self.output_button.setEnabled(True)
            self.output_button.setText(
                "Excel・PDF出力"
            )

    # =====================================
    # EstimateData作成
    # =====================================
    def create_estimate_data(
        self,
    ) -> EstimateData:

        outputs = [
            line.strip()
            for line in (
                self.outputs_edit
                .toPlainText()
                .splitlines()
            )
            if line.strip()
        ]

        return EstimateData(
            application_no=(
                self.application_edit.text().strip()
            ),
            voucher_no=(
                self.voucher_edit.text().strip()
            ),
            estimate_no=(
                self.estimate_edit.text().strip()
            ),
            issue_date=(
                self.issue_edit.text().strip()
            ),
            department=(
                self.department_edit.text().strip()
            ),
            subject=(
                self.subject_edit.text().strip()
            ),
            model_code=(
                self.model_code_edit.text().strip()
            ),
            amount=(
                self.amount_edit.text().strip()
            ),
            due_date=(
                self.delivery_edit.text().strip()
            ),
            outputs=outputs,
        )

    # =====================================
    # 終了可能か確認
    # =====================================
    def can_close(self) -> bool:

        return not (
            self.ledger_thread is not None
            and self.ledger_thread.isRunning()
        )