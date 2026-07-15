import os

from PySide6.QtWidgets import (
    QApplication,
    QFileDialog,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QMessageBox,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from services.msr_estimate_writer import MsrEstimateWriter
from services.msr_input_reader import MsrInputReader


class MsrTab(QWidget):
    """
    MSR向けの転記・台帳記入機能。

    情報を取り込むためのフォルダを選択すると、
    フォルダ内のExcelファイルをInputファイルとして自動判別し、
    フォーマットファイル（MSR用に1種類固定）をコピーした上で
    複数ファイルを一括で転記・出力する。
    """

    # 出力用フォーマット（改ざん禁止。コピーして使う）
    FORMAT_PATH = (
        "sample/"
        "【御見積書】三井E&Sシステム技研株式会社御中"
        "_準委任.xlsx"
    )

    # Inputとして扱う拡張子
    INPUT_EXTENSIONS = (".xls", ".xlsx", ".xlsm")

    def __init__(self, parent=None):
        super().__init__(parent)

        self.setup_ui()

    def setup_ui(self):

        main_layout = QVBoxLayout(self)

        # =====================================
        # タイトル
        # =====================================
        title_label = QLabel("MSR 見積書作成")
        title_label.setStyleSheet(
            "font-size: 18px; font-weight: bold;"
        )

        main_layout.addWidget(title_label)

        # =====================================
        # 情報を取り込むためのフォルダ
        # =====================================
        input_folder_label = QLabel(
            "情報を取り込むためのフォルダ"
        )

        self.input_folder_edit = QLineEdit()
        self.input_folder_edit.setPlaceholderText(
            "Inputファイルが入っているフォルダを選択してください"
        )
        self.input_folder_edit.setReadOnly(True)

        self.input_folder_button = QPushButton("参照")

        input_folder_layout = QHBoxLayout()
        input_folder_layout.addWidget(self.input_folder_edit)
        input_folder_layout.addWidget(self.input_folder_button)

        main_layout.addWidget(input_folder_label)
        main_layout.addLayout(input_folder_layout)

        # =====================================
        # 出力先フォルダ
        # =====================================
        output_folder_label = QLabel("出力先フォルダ")

        self.output_folder_edit = QLineEdit()
        self.output_folder_edit.setPlaceholderText(
            "出力先フォルダを選択してください"
        )
        self.output_folder_edit.setReadOnly(True)

        self.output_folder_button = QPushButton("参照")

        output_folder_layout = QHBoxLayout()
        output_folder_layout.addWidget(self.output_folder_edit)
        output_folder_layout.addWidget(self.output_folder_button)

        main_layout.addWidget(output_folder_label)
        main_layout.addLayout(output_folder_layout)

        # =====================================
        # 操作ボタン
        # =====================================
        self.transcribe_button = QPushButton("転記実行")
        self.ledger_write_button = QPushButton("台帳記入")

        operation_layout = QHBoxLayout()
        operation_layout.addWidget(self.transcribe_button)
        operation_layout.addWidget(self.ledger_write_button)

        main_layout.addLayout(operation_layout)

        # =====================================
        # 処理結果一覧
        # =====================================
        result_group = QGroupBox("処理結果")
        result_layout = QVBoxLayout()

        self.result_list = QListWidget()
        self.result_list.setMinimumHeight(200)

        result_layout.addWidget(self.result_list)
        result_group.setLayout(result_layout)

        main_layout.addWidget(result_group)

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
        self.input_folder_button.clicked.connect(
            self.select_input_folder
        )

        self.output_folder_button.clicked.connect(
            self.select_output_folder
        )

        self.transcribe_button.clicked.connect(
            self.transcribe
        )

        self.ledger_write_button.clicked.connect(
            self.write_ledger
        )

    # =====================================
    # 状態表示
    # =====================================
    def set_status(self, message: str):

        self.status_label.setText(message)

    # =====================================
    # 情報を取り込むためのフォルダ選択
    # =====================================
    def select_input_folder(self):

        folder_path = QFileDialog.getExistingDirectory(
            self,
            "情報を取り込むためのフォルダを選択",
        )

        if folder_path:
            self.input_folder_edit.setText(folder_path)

    # =====================================
    # 出力先フォルダ選択
    # =====================================
    def select_output_folder(self):

        folder_path = QFileDialog.getExistingDirectory(
            self,
            "出力先フォルダを選択",
        )

        if folder_path:
            self.output_folder_edit.setText(folder_path)

    # =====================================
    # 転記実行
    # =====================================
    def transcribe(self):

        input_folder = self.input_folder_edit.text().strip()
        output_folder = self.output_folder_edit.text().strip()

        if not input_folder:
            QMessageBox.warning(
                self,
                "確認",
                "情報を取り込むためのフォルダを選択してください。"
            )
            return

        if not os.path.isdir(input_folder):
            QMessageBox.warning(
                self,
                "確認",
                "情報を取り込むためのフォルダが見つかりません。\n\n"
                f"{input_folder}"
            )
            return

        if not output_folder:
            QMessageBox.warning(
                self,
                "確認",
                "出力先フォルダを選択してください。"
            )
            return

        if not os.path.isdir(output_folder):
            QMessageBox.warning(
                self,
                "確認",
                "出力先フォルダが見つかりません。\n\n"
                f"{output_folder}"
            )
            return

        if not os.path.exists(self.FORMAT_PATH):
            QMessageBox.critical(
                self,
                "エラー",
                "フォーマットファイルが見つかりません。\n\n"
                f"{os.path.abspath(self.FORMAT_PATH)}"
            )
            return

        candidates = self.find_input_candidates(
            input_folder
        )

        if not candidates:
            QMessageBox.warning(
                self,
                "確認",
                "フォルダ内にExcelファイルが見つかりません。"
            )
            return

        confirm = QMessageBox.question(
            self,
            "転記の確認",
            f"Excelファイルが {len(candidates)} 件"
            "見つかりました。\n\n"
            "見積書発行依頼として読み取れたものを"
            "すべて転記・出力します。\n"
            "実行してよろしいですか？",
            (
                QMessageBox.StandardButton.Yes
                | QMessageBox.StandardButton.No
            ),
            QMessageBox.StandardButton.No,
        )

        if confirm != QMessageBox.StandardButton.Yes:
            return

        self.run_transcription(
            candidates=candidates,
            output_folder=output_folder,
        )

    # =====================================
    # Input候補の収集
    # =====================================
    def find_input_candidates(
        self,
        input_folder: str,
    ) -> list:

        candidates = []

        for name in sorted(os.listdir(input_folder)):

            # Excelの一時ファイルは除外
            if name.startswith("~$"):
                continue

            ext = os.path.splitext(name)[1].lower()

            if ext not in self.INPUT_EXTENSIONS:
                continue

            candidates.append(
                os.path.join(input_folder, name)
            )

        return candidates

    # =====================================
    # 一括転記
    # =====================================
    def run_transcription(
        self,
        candidates: list,
        output_folder: str,
    ):

        self.result_list.clear()

        self.transcribe_button.setEnabled(False)
        self.transcribe_button.setText("転記中...")

        QApplication.processEvents()

        reader = MsrInputReader()
        writer = MsrEstimateWriter()

        success_count = 0
        skip_count = 0
        fail_count = 0

        try:

            for path in candidates:

                file_name = os.path.basename(path)

                self.set_status(
                    f"転記中：{file_name}"
                )

                QApplication.processEvents()

                # 自動判別（読み取れないものは対象外）
                try:
                    request = reader.parse(path)

                except Exception as error:
                    skip_count += 1

                    self.result_list.addItem(
                        f"対象外：{file_name}（{error}）"
                    )
                    continue

                output_name = (
                    os.path.splitext(file_name)[0]
                    + "_御見積書.xlsx"
                )

                output_path = os.path.join(
                    output_folder,
                    output_name,
                )

                try:
                    writer.write(
                        format_path=self.FORMAT_PATH,
                        output_path=output_path,
                        request=request,
                    )

                    success_count += 1

                    self.result_list.addItem(
                        f"完了：{file_name} → {output_name}"
                    )

                except Exception as error:
                    fail_count += 1

                    self.result_list.addItem(
                        f"失敗：{file_name}（{error}）"
                    )

            summary = (
                f"完了 {success_count} 件、"
                f"対象外 {skip_count} 件、"
                f"失敗 {fail_count} 件"
            )

            self.set_status(
                f"転記が終了しました。（{summary}）"
            )

            QMessageBox.information(
                self,
                "完了",
                "転記処理が終了しました。\n\n"
                f"{summary}"
            )

        finally:
            self.transcribe_button.setEnabled(True)
            self.transcribe_button.setText("転記実行")

    # =====================================
    # 管理台帳記入
    # =====================================
    def write_ledger(self):

        QMessageBox.information(
            self,
            "未実装",
            "「台帳記入」の機能は、今後ここへ実装します。"
        )
