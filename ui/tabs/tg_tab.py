import glob
import logging
import os

from PySide6.QtWidgets import (
    QApplication,
    QFileDialog,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QTextEdit,
    QVBoxLayout,
    QWidget,
    QProgressBar,
)

from services.tg_excel_writer import TgExcelWriter
from services.tg_pdf_reader import TgPdfReader

from datetime import datetime


logger = logging.getLogger(__name__)


class TgTab(QWidget):

    def __init__(self, parent=None):

        super().__init__(parent)

        self.setup_ui()

    # =====================================
    # UI
    # =====================================
    def setup_ui(self):

        layout = QVBoxLayout(self)

        # =====================================
        # タイトル
        # =====================================
        title = QLabel("TG 見積書作成")
        title.setStyleSheet(
            "font-size:18px;font-weight:bold;"
        )
        layout.addWidget(title)

        # =====================================
        # 入力フォルダ
        # =====================================
        input_label = QLabel("入力PDFフォルダ")

        self.input_dir_edit = QLineEdit()
        self.input_dir_edit.setPlaceholderText(
            "PDFが保存されているフォルダを選択してください"
        )

        self.input_dir_button = QPushButton("参照")

        input_layout = QHBoxLayout()
        input_layout.addWidget(self.input_dir_edit)
        input_layout.addWidget(self.input_dir_button)

        layout.addWidget(input_label)
        layout.addLayout(input_layout)

        # =====================================
        # 出力フォルダ
        # =====================================
        output_label = QLabel("出力フォルダ")

        self.output_dir_edit = QLineEdit()
        self.output_dir_edit.setPlaceholderText(
            "Excel・PDFの出力先フォルダ"
        )

        self.output_dir_button = QPushButton("参照")

        output_layout = QHBoxLayout()
        output_layout.addWidget(self.output_dir_edit)
        output_layout.addWidget(self.output_dir_button)

        layout.addWidget(output_label)
        layout.addLayout(output_layout)

        # =====================================
        # 操作ボタン
        # =====================================
        self.pdf_parse_button = QPushButton("PDF解析")
        self.ledger_write_button = QPushButton("台帳記入")
        self.output_button = QPushButton("Excel・PDF出力")

        button_layout = QHBoxLayout()
        button_layout.addWidget(self.pdf_parse_button)
        button_layout.addWidget(self.ledger_write_button)
        button_layout.addWidget(self.output_button)

        layout.addLayout(button_layout)

        # =====================================
        # 進捗バー
        # =====================================
        self.progress = QProgressBar()
        self.progress.setMinimum(0)
        self.progress.setValue(0)

        layout.addWidget(self.progress)

        # =====================================
        # ログ表示
        # =====================================
        log_label = QLabel("実行ログ")

        self.log_edit = QTextEdit()
        self.log_edit.setReadOnly(True)
        self.log_edit.setMinimumHeight(220)

        layout.addWidget(log_label)
        layout.addWidget(self.log_edit)

        # =====================================
        # 状態表示
        # =====================================
        self.status_label = QLabel("待機中")
        self.status_label.setStyleSheet(
            "color:#555;"
        )

        layout.addWidget(self.status_label)

        # =====================================
        # イベント
        # =====================================
        self.input_dir_button.clicked.connect(
            self.select_input_folder
        )

        self.output_dir_button.clicked.connect(
            self.select_output_folder
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
    # ステータス
    # =====================================
    def set_status(
        self,
        message: str,
        progress: int | None = None
    ):
        """
        状態表示更新
        """

        # ステータス表示
        self.status_label.setText(message)

        # ログへ出力
        self.add_log(message)

        # ProgressBar更新
        if progress is not None:
            self.progress.setValue(progress)

        # 画面更新
        QApplication.processEvents()

    # =====================================
    # ログ追加
    # =====================================
    def add_log(
        self,
        message: str
    ):
        """
        ログ表示
        """

        now = datetime.now().strftime(
            "%H:%M:%S"
        )

        self.log_edit.append(
            f"[{now}] {message}"
        )

        # 一番下へスクロール
        scrollbar = self.log_edit.verticalScrollBar()
        scrollbar.setValue(
            scrollbar.maximum()
        )

        # 画面更新
        QApplication.processEvents()

    # =====================================
    # ログクリア
    # =====================================
    def clear_log(self):

        self.log_text.clear()

        QApplication.processEvents()

    # =====================================
    # PDFフォルダ
    # =====================================
    def select_input_folder(self):
        """
        入力PDFフォルダ選択
        """

        folder = QFileDialog.getExistingDirectory(
            self,
            "PDFフォルダを選択"
        )

        if not folder:
            return

        self.input_dir_edit.setText(folder)

        # ログ表示
        self.add_log("========== 入力フォルダ ==========")
        self.add_log(folder)

        # PDF一覧表示
        pdf_files = sorted(
            glob.glob(
                os.path.join(folder, "*.pdf")
            )
        )

        if pdf_files:

            self.add_log("")
            self.add_log("読込対象PDF")

            for pdf in pdf_files:

                self.add_log(
                    f"・{os.path.basename(pdf)}"
                )

            self.add_log(
                f"合計 {len(pdf_files)} 件"
            )

        else:

            self.add_log("")
            self.add_log("PDFが見つかりません。")

    # =====================================
    # 出力フォルダ
    # =====================================
    def select_output_folder(self):
        """
        出力フォルダ選択
        """

        folder = QFileDialog.getExistingDirectory(
            self,
            "出力フォルダを選択"
        )

        if not folder:
            return

        self.output_dir_edit.setText(folder)

        self.add_log("")
        self.add_log("========== 出力フォルダ ==========")
        self.add_log(folder)

        # 既存の出力ファイルがあれば表示
        excel_files = sorted(
            glob.glob(
                os.path.join(folder, "*.xlsm")
            )
        )

        pdf_files = sorted(
            glob.glob(
                os.path.join(folder, "*.pdf")
            )
        )

        if excel_files or pdf_files:

            self.add_log("")
            self.add_log("既存ファイル")

            for file in excel_files:
                self.add_log(
                    f"Excel : {os.path.basename(file)}"
                )

            for file in pdf_files:
                self.add_log(
                    f"PDF   : {os.path.basename(file)}"
                )

        else:

            self.add_log("")
            self.add_log("出力先フォルダは空です。")


    # =====================================
    # PDFフォルダ一括解析
    # =====================================
    def pdf_parse(self):
        """
        PDFフォルダ一括解析
        """

        input_dir = self.input_dir_edit.text().strip()

        if not input_dir:

            QMessageBox.warning(
                self,
                "確認",
                "入力フォルダを選択してください。"
            )
            return

        if not os.path.isdir(input_dir):

            QMessageBox.warning(
                self,
                "確認",
                "入力フォルダが存在しません。"
            )
            return

        pdf_files = sorted(
            glob.glob(
                os.path.join(
                    input_dir,
                    "*.pdf"
                )
            )
        )

        if not pdf_files:

            QMessageBox.warning(
                self,
                "確認",
                "PDFファイルが見つかりません。"
            )
            return

        self.pdf_results = []

        self.progress.setMaximum(
            len(pdf_files)
        )
        self.progress.setValue(0)

        self.add_log("")
        self.add_log("========== PDF解析開始 ==========")
        self.add_log(f"入力フォルダ : {input_dir}")
        self.add_log(f"対象件数 : {len(pdf_files)}")

        reader = TgPdfReader()

        for index, pdf_path in enumerate(pdf_files):

            filename = os.path.basename(
                pdf_path
            )

            self.set_status(
                f"解析中 ({index + 1}/{len(pdf_files)}) {filename}",
                index
            )

            self.add_log("")
            self.add_log(
                f"----- {filename} -----"
            )

            try:

                data = reader.parse(
                    pdf_path
                )

                self.pdf_results.append(
                    {
                        "pdf_path": pdf_path,
                        "data": data,
                    }
                )

                self.add_log(
                    f"件名 : {data.subject}"
                )

                self.add_log(
                    f"金額 : {data.amount}"
                )

                self.add_log(
                    "解析完了"
                )

            except Exception as error:

                logger.exception(
                    "PDF解析失敗"
                )

                self.add_log(
                    f"解析失敗 : {error}"
                )

        self.progress.setValue(
            len(pdf_files)
        )

        self.set_status(
            "PDF解析が完了しました。",
            len(pdf_files)
        )

        self.add_log("")
        self.add_log("========== PDF解析終了 ==========")
        self.add_log(
            f"成功 : {len(self.pdf_results)} 件"
        )

        QMessageBox.information(
            self,
            "完了",
            f"{len(self.pdf_results)} 件のPDF解析が完了しました。"
        )

    # =====================================
    # 台帳記入（未実装）
    # =====================================
    def write_ledger(self):

        QMessageBox.information(
            self,
            "未実装",
            "TGの台帳記入機能は今後実装予定です。"
        )

    # =====================================
    # Excel・PDF一括出力
    # =====================================
    def export_excel(self):

        if not hasattr(self, "pdf_results") or not self.pdf_results:

            QMessageBox.warning(
                self,
                "確認",
                "先にPDF解析を実行してください。"
            )
            return

        output_dir = self.output_dir_edit.text().strip()

        if not output_dir:

            QMessageBox.warning(
                self,
                "確認",
                "出力フォルダを選択してください。"
            )
            return

        os.makedirs(
            output_dir,
            exist_ok=True
        )

        writer = TgExcelWriter()

        self.progress.setMaximum(
            len(self.pdf_results)
        )

        self.progress.setValue(0)

        self.output_button.setEnabled(False)

        self.add_log("")
        self.add_log("========== Excel・PDF出力開始 ==========")
        self.add_log(f"出力先 : {output_dir}")

        success = 0

        try:

            for index, result in enumerate(
                self.pdf_results,
                start=1
            ):

                pdf_path = result["pdf_path"]
                data = result["data"]

                filename = os.path.splitext(
                    os.path.basename(pdf_path)
                )[0]

                excel_path = os.path.join(
                    output_dir,
                    filename + ".xlsm"
                )

                pdf_output = os.path.join(
                    output_dir,
                    filename + ".pdf"
                )

                self.set_status(
                    f"出力中 ({index}/{len(self.pdf_results)})",
                    index
                )

                self.add_log("")
                self.add_log(
                    f"処理 : {filename}"
                )

                #
                # Excel作成
                #
                writer.write(
                    template_path="resources/TG_見積書フォーマット.xlsm",
                    output_path=excel_path,
                    data=data
                )

                #
                # PDF出力
                #
                writer.export_pdf(
                    excel_path,
                    pdf_output
                )

                self.add_log(
                    f"Excel : {os.path.basename(excel_path)}"
                )

                self.add_log(
                    f"PDF   : {os.path.basename(pdf_output)}"
                )

                success += 1

            self.progress.setValue(
                len(self.pdf_results)
            )

            self.set_status(
                "Excel・PDF出力完了",
                len(self.pdf_results)
            )

            self.add_log("")
            self.add_log("========== 出力終了 ==========")
            self.add_log(
                f"成功 : {success} 件"
            )

            QMessageBox.information(
                self,
                "完了",
                f"{success}件のExcel・PDFを出力しました。"
            )

        except Exception as e:

            logger.exception(
                "Excel・PDF出力失敗"
            )

            self.add_log(
                f"ERROR : {e}"
            )

            QMessageBox.critical(
                self,
                "エラー",
                str(e)
            )

        finally:

            self.output_button.setEnabled(True)