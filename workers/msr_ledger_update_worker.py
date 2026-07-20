from PySide6.QtCore import QObject, Signal, Slot

from services.msr_ledger_writer import MsrLedgerWriter
from services.onedrive_service import OneDriveService


class MsrLedgerUpdateWorker(QObject):
    """
    OneDrive上のMSR管理台帳へ、複数のInputファイル分を
    まとめて記入するバックグラウンド処理。

    台帳のダウンロード・アップロードは1回のみ行い、
    その間に各Inputファイル分の行追加を順に実行する。
    """

    finished = Signal(list)
    failed = Signal(str)
    progress = Signal(str)

    def __init__(
        self,
        share_url,
        jobs,
        device_flow_callback=None,
    ):
        """
        Args:
            share_url:
                管理台帳のOneDrive／SharePoint共有URL

            jobs:
                [{"file_name": str, "estimate_path": str,
                  "request": MsrRequest}, ...]
        """
        super().__init__()

        self.share_url = share_url
        self.jobs = jobs
        self.device_flow_callback = device_flow_callback

    @Slot()
    def run(self):
        try:
            self.progress.emit(
                "Microsoftアカウントを確認しています..."
            )

            onedrive_service = OneDriveService(
                device_flow_callback=(
                    self.device_flow_callback
                )
            )

            ledger_writer = MsrLedgerWriter()

            self.progress.emit(
                "管理台帳をダウンロードしています..."
            )

            temp_path, item = (
                onedrive_service.download_to_temp(
                    self.share_url
                )
            )

            results = []

            try:
                for job in self.jobs:

                    self.progress.emit(
                        "台帳記入中："
                        f"{job['file_name']}"
                    )

                    try:
                        result = ledger_writer.write(
                            ledger_path=temp_path,
                            estimate_path=(
                                job["estimate_path"]
                            ),
                            request=job["request"],
                        )

                        results.append({
                            "file_name": (
                                job["file_name"]
                            ),
                            "success": True,
                            "row": result["row"],
                            "estimate_no": (
                                result["estimate_no"]
                            ),
                        })

                    except Exception as error:
                        results.append({
                            "file_name": (
                                job["file_name"]
                            ),
                            "success": False,
                            "error": str(error),
                        })

                self.progress.emit(
                    "管理台帳をアップロードしています..."
                )

                onedrive_service.upload_replace(
                    local_file_path=temp_path,
                    item=item,
                )

            finally:
                import os

                if os.path.exists(temp_path):
                    try:
                        os.remove(temp_path)
                    except OSError:
                        pass

            self.progress.emit(
                "管理台帳の更新が完了しました。"
            )

            self.finished.emit(results)

        except Exception as error:
            self.failed.emit(str(error))
