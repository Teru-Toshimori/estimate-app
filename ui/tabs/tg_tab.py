import os
from collections.abc import Callable
from pathlib import Path

from PySide6.QtCore import Qt, QThread, Signal
from PySide6.QtWidgets import (
    QAbstractItemView, QHeaderView, QHBoxLayout, QLabel,
    QPushButton, QProgressBar, QTableWidget, QTableWidgetItem,
    QVBoxLayout, QWidget,
)

from ui.custom_message_dialog import (
    CustomMessageDialog,
    DialogType,
)
from ui.device_login_dialog import DeviceLoginDialog
from workers.tg_batch_worker import TgBatchWorker


class TgTab(QWidget):
    """TG見積書のPDF解析・台帳記入・Excel/PDF出力を一括実行する。"""

    RESULT_SUCCESS = "成功"
    RESULT_NG = "NG"
    RESULT_FAILED = "失敗"

    COL_REQUEST_NO = 0
    COL_ESTIMATE_NO = 1
    COL_RESULT = 2

    device_login_requested = Signal(dict)

    def __init__(self, input_provider: Callable[[], dict] | None = None, parent=None):
        super().__init__(parent)
        self.input_provider = input_provider
        self.worker_thread: QThread | None = None
        self.worker: TgBatchWorker | None = None
        self.processing = False
        self.setup_ui()
        self.device_login_requested.connect(self.show_device_login)

    def setup_ui(self) -> None:
        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(8)

        title = QLabel("TG 見積書一括作成")
        title.setStyleSheet("font-size:18px;font-weight:bold;")
        layout.addWidget(title)

        description = QLabel(
            "画面上部の共通入力欄に指定された"
            "入力フォルダ、管理台帳URL、"
            "利用者一覧URL、出力フォルダを"
            "使用して一括処理します。"
        )
        description.setWordWrap(True)
        layout.addWidget(description)

        self.execute_button = QPushButton("一括実行")
        self.execute_button.setMinimumHeight(42)
        self.cancel_button = QPushButton("処理を中止")
        self.cancel_button.setMinimumHeight(42)
        self.cancel_button.setEnabled(False)

        buttons = QHBoxLayout()
        buttons.addWidget(self.execute_button)
        buttons.addWidget(self.cancel_button)
        layout.addLayout(buttons)

        self.progress_label = QLabel("処理進捗：0 / 0件")
        layout.addWidget(self.progress_label)

        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 1)
        self.progress_bar.setValue(0)
        self.progress_bar.setFormat("%p%")
        layout.addWidget(self.progress_bar)

        self.current_file_label = QLabel("現在のファイル：なし")
        self.current_file_label.setStyleSheet("color:#555;")
        layout.addWidget(self.current_file_label)

        result_title = QLabel("処理結果")
        result_title.setStyleSheet("font-size:15px;font-weight:bold;")
        layout.addWidget(result_title)

        self.result_table = QTableWidget(0, 3)
        self.result_table.setHorizontalHeaderLabels([
            "見積依頼番号", "見積/請求番号", "結果"
        ])
        
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
        self.result_table.verticalHeader().setVisible(False)

        header = self.result_table.horizontalHeader()
        header.setSectionResizeMode(self.COL_REQUEST_NO, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(self.COL_ESTIMATE_NO, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(self.COL_RESULT, QHeaderView.ResizeMode.Fixed)
        self.result_table.setColumnWidth(self.COL_RESULT, 80)
        layout.addWidget(self.result_table, stretch=1)

        self.summary_label = QLabel("処理件数：0件　成功：0件　NG：0件　失敗：0件")
        layout.addWidget(self.summary_label)

        self.status_label = QLabel("待機中")
        self.status_label.setStyleSheet("color:#555;")
        layout.addWidget(self.status_label)

        self.execute_button.clicked.connect(self.execute_all)
        self.cancel_button.clicked.connect(self.cancel_processing)

    def get_common_inputs(self) -> dict:
        if not callable(self.input_provider):
            return {"input_folder": "", "output_folder": "", "share_url": "", "user_master_url": ""}
        values = self.input_provider() or {}
        return {
            "input_folder": str(values.get("input_folder", values.get("pdf_folder", "")) or "").strip(),
            "output_folder": str(values.get("output_folder", "") or "").strip(),
            "share_url": str(values.get("share_url", "") or "").strip(),
            "user_master_url": str(values.get("user_master_url", "") or "").strip(),
        }

    def execute_all(self) -> None:
        if self.processing:
            return

        inputs = self.get_common_inputs()
        input_folder = inputs["input_folder"]
        output_folder = inputs["output_folder"]
        share_url = inputs["share_url"]
        user_master_url = inputs["user_master_url"]

        if not self.validate_inputs(input_folder, output_folder, share_url, user_master_url):
            return

        pdf_files = sorted(Path(input_folder).glob("*.pdf"))
        if not pdf_files:
            CustomMessageDialog.warning(
                parent=self,
                title="入力内容の確認",
                heading="PDFファイルが見つかりません",
                message=(
                    "選択した入力フォルダに"
                    "PDFファイルがありません。"
                ),
            )
            return

        confirmed = CustomMessageDialog.confirm(
            parent=self,
            title="一括実行の確認",
            heading="TGの一括処理を開始します",
            message=(
                f"対象ファイル：{len(pdf_files)}件\n\n"
                "以下の処理を実行します。\n\n"
                "✓ PDF解析\n"
                "✓ 管理台帳更新\n"
                "✓ Excel・PDF出力"
            ),
            confirm_text="実行",
            cancel_text="キャンセル",
        )

        if not confirmed:
            return

        self.clear_results()
        self.set_processing_state(True)
        self.update_progress(0, len(pdf_files), "処理を準備中")
        self.status_label.setText("TG一括処理を開始します...")

        self.worker_thread = QThread(self)
        self.worker = TgBatchWorker(
            pdf_files=pdf_files,
            input_folder=input_folder,
            output_folder=output_folder,
            share_url=share_url,
            user_master_url=user_master_url,
            device_flow_callback=lambda flow: self.device_login_requested.emit(flow),
        )
        self.worker.moveToThread(self.worker_thread)
        self.worker_thread.started.connect(self.worker.run)
        self.worker.item_finished.connect(self.on_item_finished)
        self.worker.progress.connect(self.status_label.setText)
        self.worker.progress_count.connect(self.update_progress)
        self.worker.finished.connect(self.on_finished)
        self.worker.cancelled.connect(self.on_cancelled)
        self.worker.failed.connect(self.on_failed)
        self.worker.finished.connect(self.worker_thread.quit)
        self.worker.cancelled.connect(self.worker_thread.quit)
        self.worker.failed.connect(self.worker_thread.quit)
        self.worker_thread.finished.connect(self.cleanup_worker)
        self.worker_thread.start()

    def validate_inputs(
        self,
        input_folder: str,
        output_folder: str,
        share_url: str,
        user_master_url: str,
    ) -> bool:
        if not input_folder or not os.path.isdir(input_folder):
            CustomMessageDialog.warning(
                parent=self,
                title="入力内容の確認",
                heading="入力フォルダを確認してください",
                message=(
                    "共通入力の入力フォルダが未選択、"
                    "またはフォルダが見つかりません。"
                ),
            )
            return False
        if not output_folder:
            CustomMessageDialog.warning(
                parent=self,
                title="入力内容の確認",
                heading="出力フォルダが未選択です",
                message=(
                    "画面上部で出力フォルダを"
                    "選択してください。"
                ),
            )
            return False
        try:
            os.makedirs(output_folder, exist_ok=True)
        except OSError as error:
            CustomMessageDialog.error(
                parent=self,
                title="フォルダ作成エラー",
                heading="出力フォルダを作成できません",
                message=str(error),
            )
            return False
        if not share_url or not share_url.startswith(("https://", "http://")):
            CustomMessageDialog.warning(
                parent=self,
                title="入力内容の確認",
                heading="管理台帳URLを確認してください",
                message=(
                    "管理台帳URLが未入力、または"
                    "URLの形式が正しくありません。"
                ),
            )
            return False
        if not user_master_url:
            CustomMessageDialog.warning(
                parent=self,
                title="入力内容の確認",
                heading="利用者一覧URLが未入力です",
                message=(
                    "画面上部で利用者一覧URLを"
                    "入力してください。"
                ),
            )
            return False
        if not user_master_url.startswith(("https://", "http://")):
            CustomMessageDialog.warning(
                parent=self,
                title="入力内容の確認",
                heading="利用者一覧URLを確認してください",
                message="利用者一覧URLの形式が正しくありません。",
            )
            return False
        return True

    def cancel_processing(self) -> None:
        if self.worker is None:
            return

        confirmed = CustomMessageDialog.confirm(
            parent=self,
            title="処理中止の確認",
            heading="TGの一括処理を中止しますか？",
            message=(
                "現在処理中のPDFが完了したあとで"
                "処理を停止します。"
            ),
            confirm_text="中止する",
            cancel_text="処理を続ける",
        )

        if not confirmed:
            return

        self.cancel_button.setEnabled(False)
        self.cancel_button.setText("中止要求済み")
        self.status_label.setText("中止要求を受け付けました...")
        self.worker.request_cancel()

    def on_item_finished(self, result: dict) -> None:
        self.add_result_row(
            result.get("request_no", ""), result.get("estimate_no", ""),
            result.get("result", self.RESULT_FAILED), result.get("detail", ""),
        )

    def on_finished(self, summary: dict) -> None:
        self.set_processing_state(False)

        total = int(summary.get("total", 0))
        processed = int(
            summary.get(
                "processed",
                (
                    int(summary.get("success", 0))
                    + int(summary.get("ng", 0))
                    + int(summary.get("failed", 0))
                ),
            )
        )
        success = int(summary.get("success", 0))
        ng = int(summary.get("ng", 0))
        failed = int(summary.get("failed", 0))

        self.update_progress(
            processed,
            total,
            "完了",
        )
        self.status_label.setText(
            "TG一括処理が完了しました。"
        )

        CustomMessageDialog.summary(
            parent=self,
            title="処理完了",
            heading="TGの一括処理が完了しました",
            sections=[
                (
                    "TG",
                    {
                        "処理対象": total,
                        "処理済み": processed,
                        "成功": success,
                        "NG": ng,
                        "失敗": failed,
                    },
                ),
            ],
            dialog_type=DialogType.SUCCESS,
        )

    def on_cancelled(self, summary: dict) -> None:
        self.set_processing_state(False)

        total = int(summary.get("total", 0))
        processed = int(
            summary.get(
                "processed",
                (
                    int(summary.get("success", 0))
                    + int(summary.get("ng", 0))
                    + int(summary.get("failed", 0))
                ),
            )
        )
        success = int(summary.get("success", 0))
        ng = int(summary.get("ng", 0))
        failed = int(summary.get("failed", 0))

        self.update_progress(
            processed,
            total,
            "処理中止",
        )
        self.status_label.setText(
            "TG一括処理を中止しました。"
        )

        CustomMessageDialog.summary(
            parent=self,
            title="処理中止",
            heading="TGの一括処理を中止しました",
            sections=[
                (
                    "TG",
                    {
                        "処理対象": total,
                        "処理済み": processed,
                        "成功": success,
                        "NG": ng,
                        "失敗": failed,
                    },
                ),
            ],
            dialog_type=DialogType.WARNING,
        )

    def on_failed(self, message: str) -> None:
        self.set_processing_state(False)
        self.status_label.setText(
            "TG一括処理に失敗しました。"
        )

        CustomMessageDialog.error(
            parent=self,
            title="処理失敗",
            heading="TGの一括処理に失敗しました",
            message=message,
        )

    def cleanup_worker(self) -> None:
        if self.worker is not None:
            self.worker.deleteLater()
        if self.worker_thread is not None:
            self.worker_thread.deleteLater()
        self.worker = None
        self.worker_thread = None

    def set_processing_state(self, processing: bool) -> None:
        self.processing = processing
        self.execute_button.setEnabled(not processing)
        self.cancel_button.setEnabled(processing)

        if not processing:
            self.cancel_button.setText(
                "処理を中止"
            )

    def update_progress(self, current: int, total: int, file_name: str = "") -> None:
        safe_total = max(total, 1)
        self.progress_bar.setRange(0, safe_total)
        self.progress_bar.setValue(min(max(current, 0), safe_total))
        self.progress_label.setText(f"処理進捗：{current} / {total}件")
        self.current_file_label.setText(f"現在のファイル：{file_name or 'なし'}")

    def clear_results(self) -> None:
        self.result_table.setRowCount(0)
        self.update_summary()

    def add_result_row(self, request_no: str, estimate_no: str, result_text: str, detail: str = "") -> None:
        row = self.result_table.rowCount()
        self.result_table.insertRow(row)
        for column, value in enumerate(
            [request_no, estimate_no, result_text]
        ):
            item = QTableWidgetItem(
                str(value or "")
            )

            # 「結果」列のみ中央揃え
            if column == self.COL_RESULT:
                item.setTextAlignment(
                    Qt.AlignmentFlag.AlignCenter
                )

            if detail:
                item.setToolTip(detail)

            self.result_table.setItem(
                row,
                column,
                item,
            )
        self.update_summary()

    def update_summary(self) -> None:
        total = self.result_table.rowCount()
        success = ng = failed = 0
        for row in range(total):
            item = self.result_table.item(row, self.COL_RESULT)
            result = item.text().strip() if item else ""
            if result == self.RESULT_SUCCESS:
                success += 1
            elif result == self.RESULT_NG:
                ng += 1
            elif result == self.RESULT_FAILED:
                failed += 1
        self.summary_label.setText(
            f"処理件数：{total}件　成功：{success}件　NG：{ng}件　失敗：{failed}件"
        )

    def show_device_login(self, flow: dict) -> None:
        DeviceLoginDialog(flow=flow, parent=self).exec()

    def can_close(self) -> bool:
        return not (self.worker_thread is not None and self.worker_thread.isRunning())
