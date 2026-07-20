import os
import re

from PySide6.QtCore import QThread, Signal
from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QComboBox,
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
from services.msr_ledger_writer import MsrLedgerWriter
from ui.device_login_dialog import DeviceLoginDialog
from workers.msr_ledger_update_worker import (
    MsrLedgerUpdateWorker,
)


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

    device_login_requested = Signal(dict)

    def __init__(self,input_provider=None,parent=None,):
        super().__init__(parent)

        self.input_provider = input_provider

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
        title_label = QLabel("MSR 見積書作成")
        title_label.setStyleSheet(
            "font-size: 18px; font-weight: bold;"
        )

        main_layout.addWidget(title_label)

        # =====================================
        # 担当者（見積書フォーマットシート名）
        # =====================================
        staff_label = QLabel(
            "担当者（見積書フォーマットシート名）"
        )

        self.staff_combo = QComboBox()
        self.load_staff_sheets()

        main_layout.addWidget(staff_label)
        main_layout.addWidget(self.staff_combo)

        # =====================================
        # 情報を取り込むためのフォルダ
        # =====================================
        input_folder_label = QLabel(
            "見積書発行依頼の入っているフォルダ"
        )

        self.input_folder_edit = QLineEdit()
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
        self.output_folder_edit.setReadOnly(True)

        self.output_folder_button = QPushButton("参照")

        output_folder_layout = QHBoxLayout()
        output_folder_layout.addWidget(self.output_folder_edit)
        output_folder_layout.addWidget(self.output_folder_button)

        main_layout.addWidget(output_folder_label)
        main_layout.addLayout(output_folder_layout)

        # =====================================
        # 管理台帳URL（OneDrive／SharePoint）
        # =====================================
        ledger_label = QLabel(
            "管理台帳URL（OneDrive／SharePoint）"
        )

        self.ledger_edit = QLineEdit()

        self.ledger_button = QPushButton("貼り付け")

        ledger_layout = QHBoxLayout()
        ledger_layout.addWidget(self.ledger_edit)
        ledger_layout.addWidget(self.ledger_button)

        main_layout.addWidget(ledger_label)
        main_layout.addLayout(ledger_layout)

        # OneDrive未承認の間の暫定デバッグ用。
        # 承認後は削除想定。
        self.ledger_debug_checkbox = QCheckBox(
            "デバッグ用：ローカルの管理台帳ファイルを使用する"
            "（OneDrive未承認の間の暫定）"
        )

        main_layout.addWidget(
            self.ledger_debug_checkbox
        )

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

        self.ledger_button.clicked.connect(
            self.on_ledger_button_clicked
        )

        self.ledger_debug_checkbox.toggled.connect(
            self.on_ledger_debug_toggled
        )

        self.transcribe_button.clicked.connect(
            self.transcribe
        )

        self.ledger_write_button.clicked.connect(
            self.write_ledger
        )

    # =====================================
    # 担当者シート一覧の読み込み
    # =====================================
    def load_staff_sheets(self):
        """
        担当者名（フォーマットファイルのシート名）は
        フォーマット側で変わり得るため、コードに決め打ちせず
        フォーマットファイルから都度読み込む。
        """

        self.staff_combo.clear()

        try:
            sheet_names = (
                MsrEstimateWriter.list_staff_sheets(
                    self.FORMAT_PATH
                )
            )

        except Exception as error:
            QMessageBox.critical(
                self,
                "エラー",
                "フォーマットファイルのシート一覧を"
                "読み込めませんでした。\n\n"
                f"{error}"
            )
            return

        self.staff_combo.addItems(sheet_names)

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
    # デバッグ切り替え
    # =====================================
    def on_ledger_debug_toggled(self, checked: bool):

        self.ledger_edit.clear()

        if checked:
            self.ledger_button.setText("参照")
        else:
            self.ledger_button.setText("貼り付け")

    # =====================================
    # 管理台帳ボタン（URL貼り付け／ローカル参照）
    # =====================================
    def on_ledger_button_clicked(self):

        if self.ledger_debug_checkbox.isChecked():
            self.select_ledger_file()
        else:
            self.paste_ledger_url()

    # =====================================
    # 管理台帳ファイル選択（デバッグ用）
    # =====================================
    def select_ledger_file(self):

        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "管理台帳ファイルを選択",
            "",
            "Excelファイル (*.xlsx *.xlsm)"
        )

        if file_path:
            self.ledger_edit.setText(file_path)

    # =====================================
    # 管理台帳URL貼り付け
    # =====================================
    def paste_ledger_url(self):

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

        self.ledger_edit.setText(url)

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

                base_output_name = (
                    os.path.splitext(file_name)[0]
                    + "_御見積書"
                )

                output_path = self.create_unique_output_path(
                    output_folder=output_folder,
                    file_name_base=base_output_name,
                    extension=".xlsx",
                )

                output_name = os.path.basename(
                    output_path
                )

                try:
                    writer.write(
                        format_path=self.FORMAT_PATH,
                        output_path=output_path,
                        request=request,
                        staff_sheet=(
                            self.staff_combo.currentText()
                        ),
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
    # 重複しない出力パスを作成
    # =====================================
    def create_unique_output_path(
        self,
        output_folder: str,
        file_name_base: str,
        extension: str,
    ) -> str:
        """
        既存ファイルを残し、同名の場合は
        _2、_3...を付けた未使用パスを返す。
        """

        candidate = file_name_base
        number = 2

        while os.path.exists(
            os.path.join(
                output_folder,
                candidate + extension,
            )
        ):
            candidate = f"{file_name_base}_{number}"
            number += 1

        return os.path.join(
            output_folder,
            candidate + extension,
        )

    # =====================================
    # 最新の出力ファイルを取得
    # =====================================
    def find_latest_output_path(
        self,
        output_folder: str,
        file_name_base: str,
        extension: str,
    ):
        """
        元名、_2、_3...から最も大きい連番の
        見積書パスを返す。
        """

        original_path = os.path.join(
            output_folder,
            file_name_base + extension,
        )

        latest_path = (
            original_path
            if os.path.exists(original_path)
            else None
        )
        latest_number = 1 if latest_path else 0

        pattern = re.compile(
            re.escape(file_name_base)
            + r"_(\d+)"
            + re.escape(extension)
            + r"$",
            re.IGNORECASE,
        )

        try:
            names = os.listdir(output_folder)
        except OSError:
            return latest_path

        for name in names:
            match = pattern.fullmatch(name)

            if not match:
                continue

            number = int(match.group(1))

            if number > latest_number:
                latest_number = number
                latest_path = os.path.join(
                    output_folder,
                    name,
                )

        return latest_path

    # =====================================
    # 管理台帳記入
    # =====================================
    def write_ledger(self):

        input_folder = self.input_folder_edit.text().strip()
        output_folder = self.output_folder_edit.text().strip()
        ledger_value = self.ledger_edit.text().strip()
        debug_mode = (
            self.ledger_debug_checkbox.isChecked()
        )

        if not input_folder or not os.path.isdir(
            input_folder
        ):
            QMessageBox.warning(
                self,
                "確認",
                "情報を取り込むためのフォルダを選択してください。"
            )
            return

        if not output_folder or not os.path.isdir(
            output_folder
        ):
            QMessageBox.warning(
                self,
                "確認",
                "出力先フォルダを選択してください。"
            )
            return

        if debug_mode:

            if not ledger_value:
                QMessageBox.warning(
                    self,
                    "確認",
                    "管理台帳ファイルを選択してください。"
                )
                return

            if not os.path.exists(ledger_value):
                QMessageBox.warning(
                    self,
                    "確認",
                    "管理台帳ファイルが見つかりません。\n\n"
                    f"{ledger_value}"
                )
                return

        else:

            if not ledger_value:
                QMessageBox.warning(
                    self,
                    "確認",
                    "管理台帳URLを入力してください。"
                )
                return

            if not ledger_value.startswith(
                ("https://", "http://")
            ):
                QMessageBox.warning(
                    self,
                    "確認",
                    "管理台帳URLの形式が正しくありません。"
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

        # 見積書発行依頼として読み取れるものだけをジョブ化
        reader = MsrInputReader()

        jobs = []
        skip_items = []

        for path in candidates:

            file_name = os.path.basename(path)

            try:
                request = reader.parse(path)

            except Exception as error:
                skip_items.append(
                    f"対象外：{file_name}（{error}）"
                )
                continue

            base_output_name = (
                os.path.splitext(file_name)[0]
                + "_御見積書"
            )

            estimate_path = self.find_latest_output_path(
                output_folder=output_folder,
                file_name_base=base_output_name,
                extension=".xlsx",
            )

            if estimate_path is None:
                skip_items.append(
                    f"対象外：{file_name}"
                    "（出力済みの見積書が見つかりません。"
                    "先に転記実行してください）"
                )
                continue

            jobs.append({
                "file_name": file_name,
                "estimate_path": estimate_path,
                "request": request,
            })

        if not jobs:
            QMessageBox.warning(
                self,
                "確認",
                "台帳記入の対象となる見積書が"
                "見つかりませんでした。\n"
                "先に「転記実行」を行ってください。"
            )
            return

        ledger_description = (
            "ローカルの管理台帳ファイル"
            if debug_mode
            else "OneDrive／SharePoint上の管理台帳"
        )

        confirm = QMessageBox.question(
            self,
            "台帳記入の確認",
            f"{ledger_description}へ"
            f"{len(jobs)} 件分の行を追加し、"
            "対応する見積書へ見積書発行番号を"
            "転記します。\n\n"
            "実行してよろしいですか？",
            (
                QMessageBox.StandardButton.Yes
                | QMessageBox.StandardButton.No
            ),
            QMessageBox.StandardButton.No,
        )

        if confirm != QMessageBox.StandardButton.Yes:
            return

        self.result_list.clear()

        for item in skip_items:
            self.result_list.addItem(item)

        if debug_mode:
            self.run_ledger_write_local(
                ledger_path=ledger_value,
                jobs=jobs,
            )
        else:
            self.start_ledger_update(
                share_url=ledger_value,
                jobs=jobs,
            )

    # =====================================
    # ローカル台帳記入（デバッグ用・同期実行）
    # =====================================
    def run_ledger_write_local(
        self,
        ledger_path: str,
        jobs: list,
    ):

        self.ledger_write_button.setEnabled(False)
        self.ledger_write_button.setText("台帳記入中...")

        QApplication.processEvents()

        ledger_writer = MsrLedgerWriter()

        success_count = 0
        fail_count = 0

        try:

            for job in jobs:

                self.set_status(
                    f"台帳記入中：{job['file_name']}"
                )

                QApplication.processEvents()

                try:
                    result = ledger_writer.write(
                        ledger_path=ledger_path,
                        estimate_path=(
                            job["estimate_path"]
                        ),
                        request=job["request"],
                    )

                    success_count += 1

                    self.result_list.addItem(
                        f"完了：{job['file_name']} → "
                        f"台帳{result['row']}行目 "
                        f"No.{result['estimate_no']}"
                    )

                except Exception as error:
                    fail_count += 1

                    self.result_list.addItem(
                        f"失敗：{job['file_name']}"
                        f"（{error}）"
                    )

            summary = (
                f"完了 {success_count} 件、"
                f"失敗 {fail_count} 件"
            )

            self.set_status(
                f"台帳記入が終了しました。（{summary}）"
            )

            QMessageBox.information(
                self,
                "完了",
                "台帳記入処理が終了しました。\n\n"
                f"{summary}"
            )

        finally:
            self.ledger_write_button.setEnabled(True)
            self.ledger_write_button.setText("台帳記入")

    # =====================================
    # バックグラウンド台帳更新開始
    # =====================================
    def start_ledger_update(
        self,
        share_url: str,
        jobs: list,
    ):

        self.ledger_write_button.setEnabled(False)
        self.ledger_write_button.setText("台帳記入中...")

        self.set_status(
            "管理台帳の更新を開始しています..."
        )

        self.ledger_thread = QThread(self)

        self.ledger_worker = MsrLedgerUpdateWorker(
            share_url=share_url,
            jobs=jobs,
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
        results: list,
    ):

        success_count = 0
        fail_count = 0

        for result in results:

            if result["success"]:
                success_count += 1

                self.result_list.addItem(
                    f"完了：{result['file_name']} → "
                    f"台帳{result['row']}行目 "
                    f"No.{result['estimate_no']}"
                )

            else:
                fail_count += 1

                self.result_list.addItem(
                    f"失敗：{result['file_name']}"
                    f"（{result['error']}）"
                )

        summary = (
            f"完了 {success_count} 件、"
            f"失敗 {fail_count} 件"
        )

        self.set_status(
            f"台帳記入が終了しました。（{summary}）"
        )

        QMessageBox.information(
            self,
            "完了",
            "台帳記入処理が終了しました。\n\n"
            f"{summary}"
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
        self.ledger_write_button.setText("台帳記入")

        thread = self.ledger_thread

        self.ledger_worker = None
        self.ledger_thread = None

        if thread is not None:
            thread.deleteLater()

    # =====================================
    # 終了可能か確認
    # =====================================
    def can_close(self) -> bool:

        return not (
            self.ledger_thread is not None
            and self.ledger_thread.isRunning()
        )
