import os
import re

from PySide6.QtCore import Qt, QThread, Signal
from PySide6.QtWidgets import (
    QApplication,
    QAbstractItemView,
    QCheckBox,
    QComboBox,
    QFileDialog,
    QHeaderView,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QProgressBar,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from services.msr_estimate_writer import MsrEstimateWriter
from services.msr_input_reader import MsrInputReader
from services.msr_ledger_writer import MsrLedgerWriter
from services.template_resolver import TemplateResolver
from ui.device_login_dialog import DeviceLoginDialog
from workers.msr_ledger_update_worker import (
    MsrLedgerUpdateWorker,
)


class MsrTab(QWidget):
    """
    MSR向けの見積書作成・管理台帳記入機能。

    「一括実行」で次の処理を順番に行う。

    1. 入力フォルダ内のExcelを読み取る
    2. MSR見積書を出力する
    3. OneDrive／SharePoint上の管理台帳へ記入する
    4. 採番した見積/請求番号を出力見積書へ反映する

    処理結果は次の3列で表示する。

    - 見積依頼番号
    - 見積/請求番号
    - 結果（成功・NG・失敗）
    """

    INPUT_EXTENSIONS = (
        ".xls",
        ".xlsx",
        ".xlsm",
    )

    RESULT_COLUMN_REQUEST_NO = 0
    RESULT_COLUMN_ESTIMATE_NO = 1
    RESULT_COLUMN_STATUS = 2

    RESULT_SUCCESS = "成功"
    RESULT_NG = "NG"
    RESULT_FAILED = "失敗"

    device_login_requested = Signal(dict)

    def __init__(
        self,
        input_provider=None,
        parent=None,
    ):
        super().__init__(parent)

        self.input_provider = input_provider

        self.ledger_thread = None
        self.ledger_worker = None

        self.cancel_requested = False
        self.processing = False

        self.ledger_progress_current = 0
        self.ledger_progress_total = 0

        self.transcription_success_count = 0
        self.transcription_skip_count = 0
        self.transcription_fail_count = 0

        # result_keyとテーブル行番号の対応
        self.result_rows = {}

        self.setup_ui()

        self.device_login_requested.connect(
            self.show_device_login
        )

    # =====================================
    # UI
    # =====================================
    def setup_ui(self) -> None:

        main_layout = QVBoxLayout(self)

        main_layout.setContentsMargins(
            12,
            12,
            12,
            12,
        )
        main_layout.setSpacing(8)

        # タイトル
        title_label = QLabel(
            "MSR 見積書作成"
        )
        title_label.setStyleSheet(
            "font-size: 18px;"
            "font-weight: bold;"
        )
        main_layout.addWidget(title_label)

        description_label = QLabel(
            "画面上部の共通入力欄に指定された"
            "見積書発行依頼フォルダ、管理台帳URL、"
            "出力フォルダを使用し、見積書作成から"
            "管理台帳記入までを一括で処理します。"
        )
        description_label.setWordWrap(True)
        main_layout.addWidget(description_label)

        # 担当者
        staff_layout = QHBoxLayout()

        staff_label = QLabel("担当者")
        staff_label.setMinimumWidth(80)

        self.staff_combo = QComboBox()
        self.staff_combo.setMinimumHeight(32)

        staff_layout.addWidget(staff_label)
        staff_layout.addWidget(
            self.staff_combo,
            stretch=1,
        )
        main_layout.addLayout(staff_layout)

        self.load_staff_sheets()

        # =====================================
        # ▼▼▼ デバッグ用（OneDrive未承認の間の暫定）
        # OneDriveが承認されたら、この区画一式
        # （debug_local_checkbox〜local_ledger_layout）
        # と、execute_all内の分岐・
        # run_ledger_write_localメソッドを削除する。
        # =====================================
        self.debug_local_checkbox = QCheckBox(
            "デバッグ用：ローカルの管理台帳ファイルを使用する"
            "（OneDrive未承認の間の暫定）"
        )
        main_layout.addWidget(
            self.debug_local_checkbox
        )

        local_ledger_layout = QHBoxLayout()

        self.local_ledger_edit = QLineEdit()
        self.local_ledger_edit.setEnabled(False)

        self.local_ledger_button = QPushButton(
            "参照"
        )
        self.local_ledger_button.setEnabled(False)

        local_ledger_layout.addWidget(
            self.local_ledger_edit
        )
        local_ledger_layout.addWidget(
            self.local_ledger_button
        )
        main_layout.addLayout(local_ledger_layout)

        self.debug_local_checkbox.toggled.connect(
            self.on_debug_local_toggled
        )
        self.local_ledger_button.clicked.connect(
            self.select_local_ledger_file
        )
        # ▲▲▲ デバッグ用ここまで
        # =====================================

        # 操作ボタン
        self.execute_all_button = QPushButton(
            "一括実行"
        )
        self.execute_all_button.setMinimumHeight(42)

        self.cancel_button = QPushButton(
            "処理を中止"
        )
        self.cancel_button.setMinimumHeight(42)
        self.cancel_button.setEnabled(False)

        operation_layout = QHBoxLayout()
        operation_layout.addWidget(
            self.execute_all_button
        )
        operation_layout.addWidget(
            self.cancel_button
        )
        main_layout.addLayout(operation_layout)

        # 進捗表示
        self.progress_label = QLabel(
            "処理進捗：0 / 0件"
        )
        main_layout.addWidget(self.progress_label)

        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 1)
        self.progress_bar.setValue(0)
        self.progress_bar.setFormat("%p%")
        main_layout.addWidget(self.progress_bar)

        self.current_file_label = QLabel(
            "現在のファイル：なし"
        )
        self.current_file_label.setStyleSheet(
            "color: #555555;"
        )
        main_layout.addWidget(
            self.current_file_label
        )

        # 処理結果
        result_title_label = QLabel(
            "処理結果"
        )
        result_title_label.setStyleSheet(
            "font-size: 15px;"
            "font-weight: bold;"
        )
        main_layout.addWidget(
            result_title_label
        )

        self.result_table = QTableWidget(
            0,
            3,
        )
        self.result_table.setHorizontalHeaderLabels(
            [
                "見積依頼番号",
                "見積/請求番号",
                "結果",
            ]
        )
        self.result_table.setMinimumHeight(250)
        self.result_table.setEditTriggers(
            QAbstractItemView.EditTrigger.NoEditTriggers
        )
        self.result_table.setSelectionBehavior(
            QAbstractItemView.SelectionBehavior.SelectRows
        )
        self.result_table.setSelectionMode(
            QAbstractItemView.SelectionMode.SingleSelection
        )
        self.result_table.setAlternatingRowColors(False)
        self.result_table.setShowGrid(True)
        self.result_table.verticalHeader().setVisible(
            False
        )

        # 画像の特調TBに近い列幅比率
        result_header = (
            self.result_table.horizontalHeader()
        )

        result_header.setSectionResizeMode(
            self.RESULT_COLUMN_REQUEST_NO,
            QHeaderView.ResizeMode.Stretch,
        )
        result_header.setSectionResizeMode(
            self.RESULT_COLUMN_ESTIMATE_NO,
            QHeaderView.ResizeMode.Stretch,
        )
        result_header.setSectionResizeMode(
            self.RESULT_COLUMN_STATUS,
            QHeaderView.ResizeMode.Fixed,
        )

        # 「結果」列だけを細くする
        self.result_table.setColumnWidth(
            self.RESULT_COLUMN_STATUS,
            80,
        )

        main_layout.addWidget(
            self.result_table,
            stretch=1,
        )

        # 処理件数
        self.summary_label = QLabel(
            "処理件数：0件　成功：0件　NG：0件　失敗：0件"
        )
        main_layout.addWidget(
            self.summary_label
        )

        # 状態表示
        self.status_label = QLabel("待機中")
        self.status_label.setStyleSheet(
            "color: #555555;"
        )
        main_layout.addWidget(
            self.status_label
        )

        # イベント
        self.execute_all_button.clicked.connect(
            self.execute_all
        )
        self.cancel_button.clicked.connect(
            self.cancel_processing
        )

    # =====================================
    # 担当者シート
    # =====================================
    def load_staff_sheets(self) -> None:

        self.staff_combo.clear()

        try:
            template_path = TemplateResolver.resolve(
                "msr"
            )

            sheet_names = (
                MsrEstimateWriter.list_staff_sheets(
                    str(template_path)
                )
            )

        except Exception as error:
            QMessageBox.critical(
                self,
                "エラー",
                "フォーマットファイルのシート一覧を"
                "読み込めませんでした。\n\n"
                f"{error}",
            )
            return

        self.staff_combo.addItems(sheet_names)

    # =====================================
    # デバッグ用（OneDrive未承認の間の暫定）
    # OneDrive承認後、この2メソッドごと削除する。
    # =====================================
    def on_debug_local_toggled(
        self,
        checked: bool,
    ) -> None:

        self.local_ledger_edit.setEnabled(checked)
        self.local_ledger_button.setEnabled(checked)

    def select_local_ledger_file(self) -> None:

        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "管理台帳ファイルを選択",
            "",
            "Excelファイル (*.xlsx *.xlsm)",
        )

        if file_path:
            self.local_ledger_edit.setText(
                file_path
            )

    # =====================================
    # 表示制御
    # =====================================
    def update_progress(
        self,
        current: int,
        total: int,
        file_name: str = "",
    ) -> None:

        safe_total = max(total, 1)
        safe_current = min(
            max(current, 0),
            safe_total,
        )

        self.progress_bar.setRange(
            0,
            safe_total,
        )
        self.progress_bar.setValue(
            safe_current
        )

        self.progress_label.setText(
            f"処理進捗：{current} / {total}件"
        )

        self.current_file_label.setText(
            "現在のファイル："
            f"{file_name or 'なし'}"
        )

    def set_status(
        self,
        message: str,
    ) -> None:

        self.status_label.setText(message)

    def set_processing_state(
        self,
        processing: bool,
        button_text: str | None = None,
    ) -> None:

        self.processing = processing

        self.execute_all_button.setEnabled(
            not processing
        )
        self.cancel_button.setEnabled(
            processing
        )
        self.staff_combo.setEnabled(
            not processing
        )

        self.execute_all_button.setText(
            button_text
            if processing and button_text
            else "一括実行"
        )

    # =====================================
    # 結果テーブル
    # =====================================
    def clear_results(self) -> None:

        self.result_table.setRowCount(0)
        self.result_rows.clear()
        self.update_result_summary()

    def add_result_row(
        self,
        request_no: str,
        estimate_no: str = "",
        result_text: str = "",
        result_key: str | None = None,
        detail: str = "",
    ) -> int:

        row = self.result_table.rowCount()
        self.result_table.insertRow(row)

        request_item = QTableWidgetItem(
            str(request_no or "")
        )
        estimate_item = QTableWidgetItem(
            str(estimate_no or "")
        )
        result_item = QTableWidgetItem(
            str(result_text or "")
        )

        request_item.setTextAlignment(
            Qt.AlignmentFlag.AlignCenter
        )
        estimate_item.setTextAlignment(
            Qt.AlignmentFlag.AlignCenter
        )
        result_item.setTextAlignment(
            Qt.AlignmentFlag.AlignCenter
        )

        if detail:
            result_item.setToolTip(detail)

        self.result_table.setItem(
            row,
            self.RESULT_COLUMN_REQUEST_NO,
            request_item,
        )
        self.result_table.setItem(
            row,
            self.RESULT_COLUMN_ESTIMATE_NO,
            estimate_item,
        )
        self.result_table.setItem(
            row,
            self.RESULT_COLUMN_STATUS,
            result_item,
        )

        if result_key:
            self.result_rows[result_key] = row

        self.update_result_summary()

        return row

    def update_result_row(
        self,
        result_key: str,
        request_no: str | None = None,
        estimate_no: str | None = None,
        result_text: str | None = None,
        detail: str = "",
    ) -> None:

        row = self.result_rows.get(result_key)

        if row is None:
            self.add_result_row(
                request_no=request_no or "",
                estimate_no=estimate_no or "",
                result_text=result_text or "",
                result_key=result_key,
                detail=detail,
            )
            return

        if request_no is not None:
            item = self.result_table.item(
                row,
                self.RESULT_COLUMN_REQUEST_NO,
            )
            item.setText(str(request_no))

        if estimate_no is not None:
            item = self.result_table.item(
                row,
                self.RESULT_COLUMN_ESTIMATE_NO,
            )
            item.setText(str(estimate_no))

        if result_text is not None:
            item = self.result_table.item(
                row,
                self.RESULT_COLUMN_STATUS,
            )
            item.setText(str(result_text))
            item.setToolTip(detail)

        self.update_result_summary()

    def update_result_summary(self) -> None:

        total_count = self.result_table.rowCount()
        success_count = 0
        ng_count = 0
        failed_count = 0

        for row in range(total_count):
            item = self.result_table.item(
                row,
                self.RESULT_COLUMN_STATUS,
            )

            result_text = (
                item.text().strip()
                if item is not None
                else ""
            )

            if result_text == self.RESULT_SUCCESS:
                success_count += 1
            elif result_text == self.RESULT_NG:
                ng_count += 1
            elif result_text == self.RESULT_FAILED:
                failed_count += 1

        self.summary_label.setText(
            f"処理件数：{total_count}件　"
            f"成功：{success_count}件　"
            f"NG：{ng_count}件　"
            f"失敗：{failed_count}件"
        )

    # =====================================
    # 共通入力
    # =====================================
    def get_common_inputs(self) -> dict:

        if not callable(self.input_provider):
            return {
                "input_folder": "",
                "output_folder": "",
                "share_url": "",
                "user_master_url": "",
            }

        values = self.input_provider() or {}

        return {
            "input_folder": str(
                values.get(
                    "input_folder",
                    values.get(
                        "pdf_folder",
                        "",
                    ),
                )
                or ""
            ).strip(),
            "output_folder": str(
                values.get(
                    "output_folder",
                    "",
                )
                or ""
            ).strip(),
            "share_url": str(
                values.get(
                    "share_url",
                    "",
                )
                or ""
            ).strip(),
            "user_master_url": str(
                values.get(
                    "user_master_url",
                    "",
                )
                or ""
            ).strip(),
        }

    # =====================================
    # 一括実行
    # =====================================
    def execute_all(self) -> None:

        if self.processing:
            return

        if (
            self.ledger_thread is not None
            and self.ledger_thread.isRunning()
        ):
            QMessageBox.warning(
                self,
                "確認",
                "現在、処理を実行中です。",
            )
            return

        common_inputs = self.get_common_inputs()

        input_folder = common_inputs[
            "input_folder"
        ]
        output_folder = common_inputs[
            "output_folder"
        ]
        share_url = common_inputs[
            "share_url"
        ]
        user_master_url = common_inputs[
            "user_master_url"
        ]

        # デバッグ用（OneDrive未承認の間の暫定）
        debug_local = (
            self.debug_local_checkbox.isChecked()
        )
        local_ledger_path = (
            self.local_ledger_edit.text().strip()
        )

        if debug_local:
            if not self.validate_inputs_debug_local(
                input_folder=input_folder,
                output_folder=output_folder,
                local_ledger_path=local_ledger_path,
            ):
                return

        elif not self.validate_inputs(
            input_folder=input_folder,
            output_folder=output_folder,
            share_url=share_url,
            user_master_url=user_master_url,
        ):
            return

        staff_sheet = (
            self.staff_combo.currentText().strip()
        )

        if not staff_sheet:
            QMessageBox.warning(
                self,
                "確認",
                "担当者シートを選択してください。",
            )
            return

        try:
            template_path = TemplateResolver.resolve(
                "msr"
            )

        except Exception as error:
            QMessageBox.critical(
                self,
                "エラー",
                "MSR用フォーマットファイルを"
                "取得できませんでした。\n\n"
                f"{error}",
            )
            return

        candidates = self.find_input_candidates(
            input_folder
        )

        if not candidates:
            QMessageBox.warning(
                self,
                "確認",
                "フォルダ内にExcelファイルが"
                "見つかりません。",
            )
            return

        confirm = QMessageBox.question(
            self,
            "一括実行の確認",
            f"Excelファイルが {len(candidates)} 件"
            "見つかりました。\n\n"
            "見積書発行依頼として読み取れたファイルを"
            "対象に、次の処理を行います。\n\n"
            "1. 見積書の作成\n"
            "2. 管理台帳への記入\n"
            "3. 見積/請求番号の反映\n\n"
            "実行してよろしいですか？",
            (
                QMessageBox.StandardButton.Yes
                | QMessageBox.StandardButton.No
            ),
            QMessageBox.StandardButton.No,
        )

        if (
            confirm
            != QMessageBox.StandardButton.Yes
        ):
            return

        self.cancel_requested = False
        self.clear_results()

        self.transcription_success_count = 0
        self.transcription_skip_count = 0
        self.transcription_fail_count = 0

        self.set_processing_state(
            True,
            "一括処理中...",
        )

        jobs = self.run_transcription(
            candidates=candidates,
            output_folder=output_folder,
            template_path=str(template_path),
            staff_sheet=staff_sheet,
        )

        if self.cancel_requested:
            self.finish_cancelled_before_ledger()
            return

        if not jobs:
            self.finish_without_ledger_jobs()
            return

        # デバッグ用（OneDrive未承認の間の暫定）
        if debug_local:
            self.run_ledger_write_local(
                ledger_path=local_ledger_path,
                jobs=jobs,
            )
            return

        self.start_ledger_update(
            share_url=share_url,
            user_master_url=user_master_url,
            jobs=jobs,
        )

    def validate_inputs(
        self,
        input_folder: str,
        output_folder: str,
        share_url: str,
        user_master_url: str,
    ) -> bool:

        if not input_folder:
            QMessageBox.warning(
                self,
                "確認",
                "共通入力の「見積書発行依頼フォルダ」を"
                "選択してください。",
            )
            return False

        if not os.path.isdir(input_folder):
            QMessageBox.warning(
                self,
                "確認",
                "見積書発行依頼フォルダが"
                "見つかりません。\n\n"
                f"{input_folder}",
            )
            return False

        if not output_folder:
            QMessageBox.warning(
                self,
                "確認",
                "出力先フォルダを選択してください。",
            )
            return False

        if not os.path.isdir(output_folder):
            QMessageBox.warning(
                self,
                "確認",
                "出力先フォルダが"
                "見つかりません。\n\n"
                f"{output_folder}",
            )
            return False

        if not share_url:
            QMessageBox.warning(
                self,
                "確認",
                "管理台帳URLを入力してください。",
            )
            return False

        if not share_url.startswith(
            ("https://", "http://")
        ):
            QMessageBox.warning(
                self,
                "確認",
                "管理台帳URLの形式が"
                "正しくありません。",
            )
            return False

        if not user_master_url:
            QMessageBox.warning(
                self,
                "確認",
                "利用者一覧URLを入力してください。",
            )
            return False

        if not user_master_url.startswith(
            ("https://", "http://")
        ):
            QMessageBox.warning(
                self,
                "確認",
                "利用者一覧URLの形式が"
                "正しくありません。",
            )
            return False

        return True

    # =====================================
    # デバッグ用（OneDrive未承認の間の暫定）
    # OneDrive承認後、このメソッドごと削除する。
    # =====================================
    def validate_inputs_debug_local(
        self,
        input_folder: str,
        output_folder: str,
        local_ledger_path: str,
    ) -> bool:

        if not input_folder:
            QMessageBox.warning(
                self,
                "確認",
                "共通入力の「見積書発行依頼フォルダ」を"
                "選択してください。",
            )
            return False

        if not os.path.isdir(input_folder):
            QMessageBox.warning(
                self,
                "確認",
                "見積書発行依頼フォルダが"
                "見つかりません。\n\n"
                f"{input_folder}",
            )
            return False

        if not output_folder:
            QMessageBox.warning(
                self,
                "確認",
                "出力先フォルダを選択してください。",
            )
            return False

        if not os.path.isdir(output_folder):
            QMessageBox.warning(
                self,
                "確認",
                "出力先フォルダが"
                "見つかりません。\n\n"
                f"{output_folder}",
            )
            return False

        if not local_ledger_path:
            QMessageBox.warning(
                self,
                "確認",
                "デバッグ用の管理台帳ファイルを"
                "選択してください。",
            )
            return False

        if not os.path.exists(local_ledger_path):
            QMessageBox.warning(
                self,
                "確認",
                "管理台帳ファイルが"
                "見つかりません。\n\n"
                f"{local_ledger_path}",
            )
            return False

        return True

    # =====================================
    # Input候補
    # =====================================
    def find_input_candidates(
        self,
        input_folder: str,
    ) -> list[str]:

        candidates = []

        for name in sorted(
            os.listdir(input_folder)
        ):
            if name.startswith("~$"):
                continue

            extension = os.path.splitext(
                name
            )[1].lower()

            if (
                extension
                not in self.INPUT_EXTENSIONS
            ):
                continue

            candidates.append(
                os.path.join(
                    input_folder,
                    name,
                )
            )

        return candidates

    # =====================================
    # 見積書作成
    # =====================================
    def run_transcription(
        self,
        candidates: list[str],
        output_folder: str,
        template_path: str,
        staff_sheet: str,
    ) -> list[dict]:

        total = len(candidates)

        self.update_progress(
            0,
            total,
            "見積書作成を準備中",
        )
        self.set_status(
            "見積書の作成を開始します..."
        )

        reader = MsrInputReader()
        writer = MsrEstimateWriter()

        jobs = []

        for index, path in enumerate(
            candidates,
            start=1,
        ):
            QApplication.processEvents()

            if self.cancel_requested:
                break

            file_name = os.path.basename(path)

            self.update_progress(
                index,
                total,
                file_name,
            )
            self.set_status(
                "見積書作成中"
                f"（{index} / {total}件）："
                f"{file_name}"
            )
            QApplication.processEvents()

            try:
                request = reader.parse(path)

            except Exception as error:
                self.transcription_skip_count += 1

                self.add_result_row(
                    request_no="",
                    estimate_no="",
                    result_text=self.RESULT_NG,
                    detail=(
                        f"対象外：{file_name}\n"
                        f"{error}"
                    ),
                )
                continue

            # 明細（見積依頼番号）ごとに1ファイル出力
            for row in request.rows:

                request_no = row.request_no

                base_output_name = (
                    self.sanitize_file_name(
                        request_no
                    )
                )

                output_path = (
                    self.create_unique_output_path(
                        output_folder=output_folder,
                        file_name_base=(
                            base_output_name
                        ),
                        extension=".xlsx",
                    )
                )

                result_key = output_path

                self.add_result_row(
                    request_no=request_no,
                    estimate_no="",
                    result_text="",
                    result_key=result_key,
                    detail="処理中",
                )

                try:
                    writer.write(
                        format_path=template_path,
                        output_path=output_path,
                        request=request,
                        row=row,
                        staff_sheet=staff_sheet,
                    )

                    self.transcription_success_count += 1

                    jobs.append({
                        "file_name": file_name,
                        "request_no": request_no,
                        "estimate_path": output_path,
                        "request": request,
                        "row": row,
                        "result_key": result_key,
                    })

                except Exception as error:
                    self.transcription_fail_count += 1

                    self.update_result_row(
                        result_key=result_key,
                        request_no=request_no,
                        estimate_no="",
                        result_text=self.RESULT_FAILED,
                        detail=(
                            f"見積書作成失敗：{file_name}"
                            f"（{request_no}）\n"
                            f"{error}"
                        ),
                    )

        return jobs

    # =====================================
    # 出力ファイル名の無害化
    # （見積依頼番号ベース）
    # =====================================
    def sanitize_file_name(
        self,
        request_no: str,
    ) -> str:

        return re.sub(
            r'[\\/:*?"<>|]',
            "_",
            request_no,
        )

    # =====================================
    # 出力パス
    # =====================================
    def create_unique_output_path(
        self,
        output_folder: str,
        file_name_base: str,
        extension: str,
    ) -> str:

        candidate = file_name_base
        number = 2

        while os.path.exists(
            os.path.join(
                output_folder,
                candidate + extension,
            )
        ):
            candidate = (
                f"{file_name_base}_{number}"
            )
            number += 1

        return os.path.join(
            output_folder,
            candidate + extension,
        )

    # =====================================
    # 台帳更新開始
    # =====================================
    def start_ledger_update(
        self,
        share_url: str,
        user_master_url: str,
        jobs: list[dict],
    ) -> None:

        self.ledger_progress_current = 0
        self.ledger_progress_total = len(jobs)

        self.update_progress(
            0,
            self.ledger_progress_total,
            "管理台帳を準備中",
        )

        self.set_status(
            "管理台帳の更新を開始しています..."
        )

        self.execute_all_button.setText(
            "台帳更新中..."
        )

        self.ledger_thread = QThread(self)

        self.ledger_worker = (
            MsrLedgerUpdateWorker(
                share_url=share_url,
                user_master_url=user_master_url,
                jobs=jobs,
                device_flow_callback=(
                    self.request_device_login
                ),
            )
        )

        self.ledger_worker.moveToThread(
            self.ledger_thread
        )

        self.ledger_thread.started.connect(
            self.ledger_worker.run
        )

        self.ledger_worker.progress.connect(
            self.on_ledger_update_progress
        )

        self.ledger_worker.finished.connect(
            self.on_ledger_update_finished
        )

        self.ledger_worker.cancelled.connect(
            self.on_ledger_update_cancelled
        )

        self.ledger_worker.failed.connect(
            self.on_ledger_update_failed
        )

        self.ledger_worker.finished.connect(
            self.ledger_thread.quit
        )

        self.ledger_worker.cancelled.connect(
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
    # デバッグ用（OneDrive未承認の間の暫定）
    # ローカル台帳記入（同期実行）
    # OneDrive承認後、このメソッドごと削除する。
    # =====================================
    def run_ledger_write_local(
        self,
        ledger_path: str,
        jobs: list[dict],
    ) -> None:

        self.set_processing_state(
            True,
            "台帳記入中（ローカル）...",
        )

        self.update_progress(
            0,
            len(jobs),
            "管理台帳を準備中",
        )

        ledger_writer = MsrLedgerWriter()

        success_count = 0
        fail_count = 0

        try:
            for index, job in enumerate(
                jobs,
                start=1,
            ):
                QApplication.processEvents()

                self.update_progress(
                    index,
                    len(jobs),
                    job["file_name"],
                )
                self.set_status(
                    "台帳記入中（ローカル）："
                    f"{job['file_name']}"
                )
                QApplication.processEvents()

                try:
                    result = ledger_writer.write(
                        ledger_path=ledger_path,
                        estimate_path=(
                            job["estimate_path"]
                        ),
                        request=job["request"],
                        row=job["row"],
                        issuer_name="",
                    )

                    success_count += 1

                    self.update_result_row(
                        result_key=(
                            job["result_key"]
                        ),
                        request_no=(
                            job["request_no"]
                        ),
                        estimate_no=(
                            f"No.{result['estimate_no']}"
                        ),
                        result_text=(
                            self.RESULT_SUCCESS
                        ),
                        detail=(
                            f"台帳{result['row']}行目へ"
                            "記入しました。"
                        ),
                    )

                except Exception as error:
                    fail_count += 1

                    self.update_result_row(
                        result_key=(
                            job["result_key"]
                        ),
                        request_no=(
                            job["request_no"]
                        ),
                        estimate_no="",
                        result_text=(
                            self.RESULT_FAILED
                        ),
                        detail=(
                            "台帳記入に失敗しました。\n"
                            f"{error}"
                        ),
                    )

            transcription_summary = (
                self.build_transcription_summary()
            )
            ledger_summary = (
                f"台帳記入 完了 {success_count} 件、"
                f"失敗 {fail_count} 件"
            )

            self.set_status(
                "一括処理が終了しました。"
                f"（{transcription_summary}／"
                f"{ledger_summary}）"
            )

            QMessageBox.information(
                self,
                "完了",
                "一括処理が終了しました。\n\n"
                f"{transcription_summary}\n"
                f"{ledger_summary}",
            )

        finally:
            self.set_processing_state(False)

    # =====================================
    # 中止
    # =====================================
    def cancel_processing(self) -> None:

        if not self.processing:
            return

        self.cancel_requested = True
        self.cancel_button.setEnabled(False)

        if self.ledger_worker is not None:
            self.ledger_worker.request_cancel()

        self.set_status(
            "処理の中止を要求しました。"
            "現在の処理が終わり次第停止します。"
        )

        self.current_file_label.setText(
            "現在の処理：中止要求を受付済み"
        )

    def finish_cancelled_before_ledger(
        self,
    ) -> None:

        for result_key, row in self.result_rows.items():
            result_item = self.result_table.item(
                row,
                self.RESULT_COLUMN_STATUS,
            )

            if (
                result_item is not None
                and not result_item.text().strip()
            ):
                self.update_result_row(
                    result_key=result_key,
                    result_text=self.RESULT_NG,
                    detail="台帳記入前に処理を中止しました。",
                )

        summary = self.build_transcription_summary()

        self.update_progress(
            0,
            0,
            "中止",
        )
        self.set_status(
            "見積書作成処理を中止しました。"
            f"（{summary}）"
        )

        self.set_processing_state(False)

        QMessageBox.information(
            self,
            "処理中止",
            "一括処理を中止しました。\n\n"
            f"{summary}",
        )

    def finish_without_ledger_jobs(
        self,
    ) -> None:

        summary = self.build_transcription_summary()

        self.set_status(
            "管理台帳へ記入できる見積書が"
            "作成されませんでした。"
            f"（{summary}）"
        )

        self.set_processing_state(False)

        QMessageBox.warning(
            self,
            "処理終了",
            "管理台帳へ記入できる見積書が"
            "作成されませんでした。\n\n"
            f"{summary}",
        )

    # =====================================
    # 台帳更新イベント
    # =====================================
    def on_ledger_update_progress(
        self,
        message: str,
    ) -> None:

        self.set_status(message)

        prefix = "台帳記入中："

        if message.startswith(prefix):
            self.ledger_progress_current = min(
                self.ledger_progress_current + 1,
                self.ledger_progress_total,
            )

            file_name = message[
                len(prefix):
            ].strip()

            self.update_progress(
                self.ledger_progress_current,
                self.ledger_progress_total,
                file_name,
            )
            return

        if "アップロード" in message:
            display_text = (
                "管理台帳をアップロード中"
            )
        elif "ダウンロード" in message:
            display_text = (
                "管理台帳をダウンロード中"
            )
        elif "Microsoftアカウント" in message:
            display_text = (
                "Microsoftアカウントを確認中"
            )
        else:
            display_text = message

        self.current_file_label.setText(
            f"現在の処理：{display_text}"
        )

    def on_ledger_update_finished(
        self,
        results: list,
    ) -> None:

        success_count = 0
        fail_count = 0

        for result in results:
            result_key = result[
                "result_key"
            ]
            request_no = result[
                "request_no"
            ]

            if result["success"]:
                success_count += 1

                estimate_no = (
                    f"No.{result['estimate_no']}"
                )

                self.update_result_row(
                    result_key=result_key,
                    request_no=request_no,
                    estimate_no=estimate_no,
                    result_text=self.RESULT_SUCCESS,
                    detail=(
                        f"台帳{result['row']}行目へ"
                        "記入しました。"
                    ),
                )

            else:
                fail_count += 1

                self.update_result_row(
                    result_key=result_key,
                    request_no=request_no,
                    estimate_no="",
                    result_text=self.RESULT_FAILED,
                    detail=(
                        "台帳記入に失敗しました。\n"
                        f"{result['error']}"
                    ),
                )

        self.update_progress(
            self.ledger_progress_total,
            self.ledger_progress_total,
            "完了",
        )

        transcription_summary = (
            self.build_transcription_summary()
        )
        ledger_summary = (
            f"台帳記入 完了 {success_count} 件、"
            f"失敗 {fail_count} 件"
        )

        self.set_status(
            "一括処理が終了しました。"
            f"（{transcription_summary}／"
            f"{ledger_summary}）"
        )

        QMessageBox.information(
            self,
            "完了",
            "一括処理が終了しました。\n\n"
            f"{transcription_summary}\n"
            f"{ledger_summary}",
        )

    def on_ledger_update_cancelled(
        self,
        results: list,
    ) -> None:

        completed_keys = set()
        success_count = 0
        fail_count = 0

        for result in results:
            result_key = result[
                "result_key"
            ]
            completed_keys.add(result_key)

            request_no = result[
                "request_no"
            ]

            if result["success"]:
                success_count += 1

                self.update_result_row(
                    result_key=result_key,
                    request_no=request_no,
                    estimate_no=(
                        f"No.{result['estimate_no']}"
                    ),
                    result_text=self.RESULT_SUCCESS,
                    detail=(
                        f"台帳{result['row']}行目へ"
                        "記入しました。"
                    ),
                )

            else:
                fail_count += 1

                self.update_result_row(
                    result_key=result_key,
                    request_no=request_no,
                    estimate_no="",
                    result_text=self.RESULT_FAILED,
                    detail=(
                        "台帳記入に失敗しました。\n"
                        f"{result['error']}"
                    ),
                )

        for result_key, row in self.result_rows.items():
            if result_key in completed_keys:
                continue

            result_item = self.result_table.item(
                row,
                self.RESULT_COLUMN_STATUS,
            )

            if (
                result_item is not None
                and not result_item.text().strip()
            ):
                self.update_result_row(
                    result_key=result_key,
                    result_text=self.RESULT_NG,
                    detail="台帳記入前に処理を中止しました。",
                )

        summary = (
            f"中止前に台帳記入 完了 "
            f"{success_count} 件、"
            f"失敗 {fail_count} 件"
        )

        self.set_status(
            "一括処理を中止しました。"
            f"（{summary}）"
        )

        self.current_file_label.setText(
            "現在の処理：中止"
        )

        QMessageBox.information(
            self,
            "処理中止",
            "一括処理を中止しました。\n\n"
            f"{self.build_transcription_summary()}\n"
            f"{summary}",
        )

    def on_ledger_update_failed(
        self,
        error_message: str,
    ) -> None:

        for result_key, row in self.result_rows.items():
            result_item = self.result_table.item(
                row,
                self.RESULT_COLUMN_STATUS,
            )

            if (
                result_item is not None
                and not result_item.text().strip()
            ):
                self.update_result_row(
                    result_key=result_key,
                    result_text=self.RESULT_FAILED,
                    detail=(
                        "管理台帳の更新に失敗しました。\n"
                        f"{error_message}"
                    ),
                )

        self.set_status(
            "管理台帳の更新に失敗しました。"
        )
        self.current_file_label.setText(
            "現在の処理：エラー"
        )

        QMessageBox.critical(
            self,
            "エラー",
            "OneDrive／SharePointの台帳記入に"
            "失敗しました。\n\n"
            f"{error_message}",
        )

    # =====================================
    # 認証
    # =====================================
    def request_device_login(
        self,
        flow: dict,
    ) -> None:

        self.device_login_requested.emit(flow)

    def show_device_login(
        self,
        flow: dict,
    ) -> None:

        dialog = DeviceLoginDialog(
            flow=flow,
            parent=self,
        )

        dialog.exec()

    # =====================================
    # 後処理
    # =====================================
    def clear_ledger_thread(self) -> None:

        thread = self.ledger_thread

        self.ledger_worker = None
        self.ledger_thread = None

        self.cancel_requested = False
        self.set_processing_state(False)

        if thread is not None:
            thread.deleteLater()

    def build_transcription_summary(self) -> str:

        return (
            "見積書作成 "
            f"完了 {self.transcription_success_count} 件、"
            f"対象外 {self.transcription_skip_count} 件、"
            f"失敗 {self.transcription_fail_count} 件"
        )

    def can_close(self) -> bool:

        return not self.processing
