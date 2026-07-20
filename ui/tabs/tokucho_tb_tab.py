from collections.abc import Callable
from pathlib import Path

from PySide6.QtCore import (
    QThread,
    Qt,
)
from PySide6.QtWidgets import (
    QHeaderView,
    QHBoxLayout,
    QLabel,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from ui.device_login_dialog import DeviceLoginDialog
from workers.batch_estimate_worker import BatchEstimateWorker


class TokuchoTbTab(QWidget):
    """
    特調TB向けの見積書一括作成タブ。

    共通入力値はMainWindowから取得する。
    """

    def __init__(
        self,
        input_provider: Callable[[], dict],
        parent=None,
    ):
        super().__init__(parent)

        self.input_provider = input_provider

        self.batch_thread: QThread | None = None
        self.batch_worker: BatchEstimateWorker | None = None

        self.result_details: list[dict] = []

        self.total_count = 0
        self.success_count = 0
        self.ng_count = 0
        self.failed_count = 0

        self.setup_ui()

    # =====================================
    # 画面作成
    # =====================================
    def setup_ui(self) -> None:

        main_layout = QVBoxLayout(self)

        main_layout.setContentsMargins(
            12,
            12,
            12,
            12,
        )

        main_layout.setSpacing(
            8
        )

        title_label = QLabel(
            "特調TB 見積書一括作成"
        )

        title_label.setStyleSheet(
            "font-size: 18px;"
            "font-weight: bold;"
        )

        main_layout.addWidget(
            title_label
        )

        description_label = QLabel(
            "画面上部の共通入力欄に指定された"
            "業務計画書フォルダ、管理台帳URL、"
            "利用者一覧URL、出力フォルダを"
            "使用して一括処理します。"
        )

        description_label.setWordWrap(
            True
        )

        main_layout.addWidget(
            description_label
        )

        # =====================================
        # 実行・中止ボタン
        # =====================================
        self.batch_execute_button = QPushButton(
            "一括実行"
        )

        self.batch_execute_button.setMinimumHeight(
            42
        )

        self.cancel_button = QPushButton(
            "処理を中止"
        )

        self.cancel_button.setMinimumHeight(
            42
        )

        self.cancel_button.setEnabled(
            False
        )

        button_layout = QHBoxLayout()

        button_layout.addWidget(
            self.batch_execute_button
        )

        button_layout.addWidget(
            self.cancel_button
        )

        main_layout.addLayout(
            button_layout
        )

        # =====================================
        # 進捗表示
        # =====================================
        self.progress_label = QLabel(
            "処理進捗：0 / 0件"
        )

        main_layout.addWidget(
            self.progress_label
        )

        self.progress_bar = QProgressBar()

        self.progress_bar.setRange(
            0,
            1,
        )

        self.progress_bar.setValue(
            0
        )

        self.progress_bar.setFormat(
            "%p%"
        )

        main_layout.addWidget(
            self.progress_bar
        )

        self.current_file_label = QLabel(
            "現在のファイル：なし"
        )

        self.current_file_label.setStyleSheet(
            "color: #555555;"
        )

        main_layout.addWidget(
            self.current_file_label
        )

        # =====================================
        # 処理結果
        # =====================================
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
                "伝票番号",
                "見積/請求番号",
                "結果",
            ]
        )

        self.result_table.setEditTriggers(
            QTableWidget.EditTrigger.NoEditTriggers
        )

        self.result_table.setSelectionBehavior(
            QTableWidget.SelectionBehavior.SelectRows
        )

        self.result_table.setSelectionMode(
            QTableWidget.SelectionMode.SingleSelection
        )

        self.result_table.verticalHeader().setVisible(
            False
        )

        table_header = (
            self.result_table.horizontalHeader()
        )

        table_header.setSectionResizeMode(
            0,
            QHeaderView.ResizeMode.Stretch,
        )

        table_header.setSectionResizeMode(
            1,
            QHeaderView.ResizeMode.Stretch,
        )

        table_header.setSectionResizeMode(
            2,
            QHeaderView.ResizeMode.ResizeToContents,
        )

        main_layout.addWidget(
            self.result_table,
            stretch=1,
        )

        self.summary_label = QLabel(
            "処理件数：0件　"
            "成功：0件　"
            "NG：0件　"
            "失敗：0件"
        )

        main_layout.addWidget(
            self.summary_label
        )

        self.status_label = QLabel(
            "待機中"
        )

        self.status_label.setStyleSheet(
            "color: #555555;"
        )

        main_layout.addWidget(
            self.status_label
        )

        self.batch_execute_button.clicked.connect(
            self.start_batch_process
        )

        self.cancel_button.clicked.connect(
            self.cancel_batch_process
        )

        self.result_table.cellDoubleClicked.connect(
            self.show_result_detail
        )

    # =====================================
    # 一括処理開始
    # =====================================
    def start_batch_process(self) -> None:

        if (
            self.batch_thread is not None
            and self.batch_thread.isRunning()
        ):
            QMessageBox.warning(
                self,
                "確認",
                "現在、一括処理を実行中です。",
            )
            return

        inputs = self.input_provider()

        pdf_folder = str(
            inputs.get(
                "pdf_folder",
                "",
            )
        ).strip()

        share_url = str(
            inputs.get(
                "share_url",
                "",
            )
        ).strip()

        user_master_url = str(
            inputs.get(
                "user_master_url",
                "",
            )
        ).strip()

        output_folder = str(
            inputs.get(
                "output_folder",
                "",
            )
        ).strip()

        if not self.validate_common_inputs(
            pdf_folder=pdf_folder,
            share_url=share_url,
            user_master_url=user_master_url,
            output_folder=output_folder,
        ):
            return

        pdf_files = self.get_pdf_files(
            pdf_folder
        )

        if not pdf_files:
            QMessageBox.warning(
                self,
                "確認",
                "選択した業務計画書フォルダに"
                "PDFファイルがありません。",
            )
            return

        confirmation = QMessageBox.question(
            self,
            "一括実行の確認",
            f"{len(pdf_files)}件のPDFを処理します。\n\n"
            f"業務計画書フォルダ：\n{pdf_folder}\n\n"
            f"管理台帳URL：\n{share_url}\n\n"
            f"利用者一覧URL：\n{user_master_url}\n\n"
            f"出力フォルダ：\n{output_folder}\n\n"
            "実行してよろしいですか？",
            (
                QMessageBox.StandardButton.Yes
                | QMessageBox.StandardButton.No
            ),
            QMessageBox.StandardButton.No,
        )

        if (
            confirmation
            != QMessageBox.StandardButton.Yes
        ):
            return

        self.clear_results()

        self.start_batch_worker(
            pdf_files=pdf_files,
            share_url=share_url,
            user_master_url=user_master_url,
            output_folder=output_folder,
        )

    # =====================================
    # 共通入力チェック
    # =====================================
    def validate_common_inputs(
        self,
        pdf_folder: str,
        share_url: str,
        user_master_url: str,
        output_folder: str,
    ) -> bool:

        if not pdf_folder:
            QMessageBox.warning(
                self,
                "確認",
                "画面上部で業務計画書フォルダを"
                "選択してください。",
            )
            return False

        pdf_folder_path = Path(
            pdf_folder
        )

        if not pdf_folder_path.is_dir():
            QMessageBox.warning(
                self,
                "確認",
                "業務計画書フォルダが"
                "見つかりません。\n\n"
                f"{pdf_folder}",
            )
            return False

        if not self.is_valid_url(
            share_url
        ):
            QMessageBox.warning(
                self,
                "確認",
                "管理台帳URLを正しく"
                "入力してください。",
            )
            return False

        if not self.is_valid_url(
            user_master_url
        ):
            QMessageBox.warning(
                self,
                "確認",
                "利用者一覧URLを正しく"
                "入力してください。",
            )
            return False

        if not output_folder:
            QMessageBox.warning(
                self,
                "確認",
                "画面上部で出力フォルダを"
                "選択してください。",
            )
            return False

        output_folder_path = Path(
            output_folder
        )

        if not output_folder_path.is_dir():
            QMessageBox.warning(
                self,
                "確認",
                "出力フォルダが"
                "見つかりません。\n\n"
                f"{output_folder}",
            )
            return False

        return True

    # =====================================
    # URLチェック
    # =====================================
    def is_valid_url(
        self,
        value: str,
    ) -> bool:

        return bool(
            value
            and value.startswith(
                (
                    "https://",
                    "http://",
                )
            )
        )

    # =====================================
    # Worker開始
    # =====================================
    def start_batch_worker(
        self,
        pdf_files: list[Path],
        share_url: str,
        user_master_url: str,
        output_folder: str,
    ) -> None:

        self.update_progress(
            current=0,
            total=len(pdf_files),
            file_name="",
        )

        self.batch_execute_button.setEnabled(
            False
        )

        self.batch_execute_button.setText(
            "一括処理中..."
        )

        self.cancel_button.setEnabled(
            True
        )

        self.set_status(
            "一括処理を開始しています..."
        )

        self.batch_thread = QThread(
            self
        )

        self.batch_worker = BatchEstimateWorker(
            pdf_files=pdf_files,
            share_url=share_url,
            user_master_url=user_master_url,
            output_folder=output_folder,
        )

        self.batch_worker.moveToThread(
            self.batch_thread
        )

        self.batch_thread.started.connect(
            self.batch_worker.run
        )

        self.batch_worker.progress.connect(
            self.set_status
        )

        self.batch_worker.progress_count.connect(
            self.update_progress
        )

        self.batch_worker.device_login_requested.connect(
            self.show_device_login
        )

        self.batch_worker.item_finished.connect(
            self.on_batch_item_finished
        )

        self.batch_worker.batch_finished.connect(
            self.on_batch_finished
        )

        self.batch_worker.batch_failed.connect(
            self.on_batch_failed
        )

        self.batch_worker.batch_finished.connect(
            self.batch_thread.quit
        )

        self.batch_worker.batch_failed.connect(
            self.batch_thread.quit
        )

        self.batch_thread.finished.connect(
            self.batch_worker.deleteLater
        )

        self.batch_thread.finished.connect(
            self.clear_batch_thread
        )

        self.batch_thread.start()

    # =====================================
    # 中止
    # =====================================
    def cancel_batch_process(self) -> None:

        if self.batch_worker is None:
            return

        confirmation = QMessageBox.question(
            self,
            "処理中止の確認",
            "一括処理を中止しますか？\n\n"
            "現在処理中のPDFが完了したあとで"
            "停止します。",
            (
                QMessageBox.StandardButton.Yes
                | QMessageBox.StandardButton.No
            ),
            QMessageBox.StandardButton.No,
        )

        if (
            confirmation
            != QMessageBox.StandardButton.Yes
        ):
            return

        self.batch_worker.cancel()

        self.cancel_button.setEnabled(
            False
        )

        self.cancel_button.setText(
            "中止要求済み"
        )

        self.set_status(
            "処理の中止を要求しました。"
        )

    # =====================================
    # Microsoft認証
    # =====================================
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
    # 結果受信
    # =====================================
    def on_batch_item_finished(
        self,
        result: dict,
    ) -> None:

        self.result_details.append(
            result
        )

        result_text = str(
            result.get(
                "result",
                "",
            )
        )

        self.add_result_row(
            voucher_no=str(
                result.get(
                    "voucher_no",
                    "",
                )
            ),
            estimate_no=str(
                result.get(
                    "estimate_no",
                    "",
                )
            ),
            result_text=result_text,
        )

        self.total_count += 1

        if result_text == "成功":
            self.success_count += 1

        elif result_text == "NG":
            self.ng_count += 1

        else:
            self.failed_count += 1

        self.update_summary(
            total=self.total_count,
            success=self.success_count,
            ng=self.ng_count,
            failed=self.failed_count,
        )

    # =====================================
    # 完了
    # =====================================
    def on_batch_finished(
        self,
        summary: dict,
    ) -> None:

        total = int(
            summary.get(
                "total",
                0,
            )
        )

        processed = int(
            summary.get(
                "processed",
                0,
            )
        )

        success = int(
            summary.get(
                "success",
                0,
            )
        )

        ng = int(
            summary.get(
                "ng",
                0,
            )
        )

        failed = int(
            summary.get(
                "failed",
                0,
            )
        )

        cancelled = bool(
            summary.get(
                "cancelled",
                False,
            )
        )

        self.update_progress(
            current=processed,
            total=total,
            file_name="",
        )

        self.update_summary(
            total=processed,
            success=success,
            ng=ng,
            failed=failed,
        )

        if cancelled:
            title = "中止"
            self.set_status(
                "一括処理を中止しました。"
            )
        else:
            title = "完了"
            self.set_status(
                "一括処理が完了しました。"
            )

        QMessageBox.information(
            self,
            title,
            f"処理対象：{total}件\n"
            f"処理済み：{processed}件\n"
            f"成功：{success}件\n"
            f"NG：{ng}件\n"
            f"失敗：{failed}件",
        )

    # =====================================
    # 全体エラー
    # =====================================
    def on_batch_failed(
        self,
        error_message: str,
    ) -> None:

        self.set_status(
            "一括処理に失敗しました。"
        )

        self.current_file_label.setText(
            "現在のファイル：処理停止"
        )

        QMessageBox.critical(
            self,
            "エラー",
            error_message,
        )

    # =====================================
    # 詳細表示
    # =====================================
    def show_result_detail(
        self,
        row: int,
        column: int,
    ) -> None:

        _ = column

        if not (
            0
            <= row
            < len(self.result_details)
        ):
            return

        result = self.result_details[row]

        detail_message = (
            f"伝票番号："
            f"{result.get('voucher_no', '')}\n"
            f"見積/請求番号："
            f"{result.get('estimate_no', '') or '未採番'}\n"
            f"結果："
            f"{result.get('result', '')}\n\n"
            f"詳細：\n"
            f"{result.get('message', '') or '詳細情報なし'}"
        )

        QMessageBox.information(
            self,
            "処理結果詳細",
            detail_message,
        )

    # =====================================
    # 進捗
    # =====================================
    def update_progress(
        self,
        current: int,
        total: int,
        file_name: str,
    ) -> None:

        safe_total = max(
            total,
            1,
        )

        safe_current = min(
            max(
                current,
                0,
            ),
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
            f"現在のファイル："
            f"{file_name or 'なし'}"
        )

    # =====================================
    # 状態表示
    # =====================================
    def set_status(
        self,
        message: str,
    ) -> None:

        self.status_label.setText(
            message
        )

    # =====================================
    # PDF一覧
    # =====================================
    def get_pdf_files(
        self,
        folder_path: str,
    ) -> list[Path]:

        folder = Path(
            folder_path
        )

        if not folder.is_dir():
            return []

        return sorted(
            [
                path
                for path in folder.iterdir()
                if (
                    path.is_file()
                    and path.suffix.lower()
                    == ".pdf"
                )
            ],
            key=lambda path: (
                path.name.casefold()
            ),
        )

    # =====================================
    # 結果初期化
    # =====================================
    def clear_results(self) -> None:

        self.result_table.setRowCount(
            0
        )

        self.result_details.clear()

        self.total_count = 0
        self.success_count = 0
        self.ng_count = 0
        self.failed_count = 0

        self.update_summary(
            total=0,
            success=0,
            ng=0,
            failed=0,
        )

        self.update_progress(
            current=0,
            total=0,
            file_name="",
        )

    # =====================================
    # 結果行追加
    # =====================================
    def add_result_row(
        self,
        voucher_no: str,
        estimate_no: str,
        result_text: str,
    ) -> None:

        row = self.result_table.rowCount()

        self.result_table.insertRow(
            row
        )

        voucher_item = QTableWidgetItem(
            voucher_no
        )

        estimate_item = QTableWidgetItem(
            estimate_no
        )

        result_item = QTableWidgetItem(
            result_text
        )

        result_item.setTextAlignment(
            Qt.AlignmentFlag.AlignCenter
        )

        self.result_table.setItem(
            row,
            0,
            voucher_item,
        )

        self.result_table.setItem(
            row,
            1,
            estimate_item,
        )

        self.result_table.setItem(
            row,
            2,
            result_item,
        )

    # =====================================
    # 件数表示
    # =====================================
    def update_summary(
        self,
        total: int,
        success: int,
        ng: int,
        failed: int,
    ) -> None:

        self.summary_label.setText(
            f"処理件数：{total}件　"
            f"成功：{success}件　"
            f"NG：{ng}件　"
            f"失敗：{failed}件"
        )

    # =====================================
    # スレッド後始末
    # =====================================
    def clear_batch_thread(self) -> None:

        self.batch_execute_button.setEnabled(
            True
        )

        self.batch_execute_button.setText(
            "一括実行"
        )

        self.cancel_button.setEnabled(
            False
        )

        self.cancel_button.setText(
            "処理を中止"
        )

        thread = self.batch_thread

        self.batch_worker = None
        self.batch_thread = None

        if thread is not None:
            thread.deleteLater()

    # =====================================
    # 終了可能判定
    # =====================================
    def can_close(self) -> bool:

        return not (
            self.batch_thread is not None
            and self.batch_thread.isRunning()
        )