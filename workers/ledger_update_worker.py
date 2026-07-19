from PySide6.QtCore import QObject, Signal, Slot

from services.ledger_writer import LedgerWriter
from services.onedrive_service import OneDriveService


class LedgerUpdateWorker(QObject):
    """OneDrive上の管理台帳を更新するバックグラウンド処理。"""

    finished = Signal(dict)
    failed = Signal(str)
    progress = Signal(str)

    def __init__(
        self,
        share_url,
        data,
        device_flow_callback=None,
    ):
        super().__init__()

        self.share_url = share_url
        self.data = data
        self.device_flow_callback = device_flow_callback

    @Slot()
    def run(self):
        try:
            self.progress.emit(
                "Microsoftアカウントを確認しています..."
            )

            onedrive_service = OneDriveService(
                device_flow_callback=self.device_flow_callback
            )

            ledger_writer = LedgerWriter()

            self.progress.emit(
                "管理台帳をダウンロードしています..."
            )

            result = onedrive_service.update_ledger(
                share_url=self.share_url,
                ledger_writer=ledger_writer,
                data=self.data,
            )

            self.progress.emit(
                "管理台帳の更新が完了しました。"
            )

            self.finished.emit(result)

        except Exception as error:
            self.failed.emit(str(error))