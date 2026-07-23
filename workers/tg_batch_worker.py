import os
import re
import threading
from pathlib import Path

from PySide6.QtCore import QObject, Signal, Slot

from services.cloud.onedrive_service import OneDriveService
from services.writers.tg_excel_writer import TgExcelWriter
from services.writers.tg_ledger_writer import (
    DuplicateRequestNoError,
    TgLedgerWriter,
)
from services.readers.tg_pdf_reader import TgPdfReader


class TgBatchWorker(QObject):
    item_finished = Signal(dict)
    progress = Signal(str)
    progress_count = Signal(int, int, str)
    finished = Signal(dict)
    cancelled = Signal(dict)
    failed = Signal(str)

    def __init__(
        self,
        pdf_files,
        input_folder,
        output_folder,
        share_url,
        user_master_url,
        device_flow_callback=None,
    ):
        super().__init__()

        self.pdf_files = [Path(path) for path in pdf_files]
        self.input_folder = Path(input_folder)
        self.output_folder = Path(output_folder)
        self.share_url = str(share_url or "").strip()
        self.user_master_url = str(user_master_url or "").strip()
        self.device_flow_callback = device_flow_callback
        self._cancel_event = threading.Event()

    def request_cancel(self) -> None:
        self._cancel_event.set()

    def is_cancel_requested(self) -> bool:
        return self._cancel_event.is_set()

    @Slot()
    def run(self) -> None:
        total = len(self.pdf_files)
        success = 0
        ng = 0
        failed = 0

        temp_ledger_path = None
        ledger_item = None
        ledger_changed = False

        try:
            self.output_folder.mkdir(parents=True, exist_ok=True)

            self.progress.emit(
                "Microsoftアカウントを確認しています..."
            )

            onedrive = OneDriveService(
                device_flow_callback=self.device_flow_callback,
                user_master_url=self.user_master_url,
            )

            self.progress.emit(
                "利用者情報を取得しています..."
            )
            issuer_name = onedrive.get_issuer_name_from_master()

            if self.is_cancel_requested():
                self.cancelled.emit({
                    "total": total,
                    "processed": (
                        success + ng + failed
                    ),
                    "success": success,
                    "ng": ng,
                    "failed": failed,
                    "cancelled": True,
                })
                return

            self.progress.emit(
                "管理台帳をダウンロードしています..."
            )
            temp_ledger_path, ledger_item = (
                onedrive.download_to_temp(self.share_url)
            )

            pdf_reader = TgPdfReader()
            ledger_writer = TgLedgerWriter()
            excel_writer = TgExcelWriter()
            excel_files = self.find_excel_files()

            for index, pdf_path in enumerate(
                self.pdf_files,
                start=1,
            ):
                if self.is_cancel_requested():
                    break

                request_no = self.extract_request_no(pdf_path)

                self.progress_count.emit(
                    index,
                    total,
                    pdf_path.name,
                )
                self.progress.emit(
                    f"TG処理中（{index} / {total}件）："
                    f"{pdf_path.name}"
                )

                excel_path = self.find_excel_for_pdf(
                    pdf_path,
                    excel_files,
                )

                if excel_path is None:
                    ng += 1
                    self.item_finished.emit({
                        "request_no": request_no,
                        "estimate_no": "",
                        "result": "NG",
                        "detail": (
                            "PDFに対応する元Excel（xlsm）が"
                            "見つかりません。"
                        ),
                    })
                    continue

                try:
                    data = pdf_reader.parse(str(pdf_path))

                    subject = str(
                        data.get("品名", "") or ""
                    ).strip()
                    amount = str(
                        data.get("金額", "") or ""
                    ).strip()

                    if not subject or not amount:
                        ng += 1
                        self.item_finished.emit({
                            "request_no": request_no,
                            "estimate_no": "",
                            "result": "NG",
                            "detail": (
                                "PDFから品名または金額を"
                                "取得できませんでした。"
                            ),
                        })
                        continue

                    ledger_result = ledger_writer.append(
                        ledger_path=temp_ledger_path,
                        request_no=request_no,
                        data=data,
                        issuer_name=issuer_name,
                    )

                    estimate_no = str(
                        ledger_result["estimate_no"]
                    )
                    ledger_changed = True

                    output_excel = self.create_unique_output_path(
                        pdf_path.stem,
                        ".xlsm",
                    )
                    output_pdf = output_excel.with_suffix(".pdf")

                    excel_writer.write(
                        str(excel_path),
                        str(output_excel),
                        data,
                        estimate_no,
                    )
                    excel_writer.export_pdf(
                        str(output_excel),
                        str(output_pdf),
                    )

                    success += 1
                    self.item_finished.emit({
                        "request_no": request_no,
                        "estimate_no": estimate_no,
                        "result": "成功",
                        "detail": (
                            f"Excel: {output_excel.name}\n"
                            f"PDF: {output_pdf.name}\n"
                            f"発行者: {issuer_name}"
                        ),
                    })

                except DuplicateRequestNoError as error:
                    # 重複した案件だけNGにして、次のPDFへ進む。
                    ng += 1
                    self.item_finished.emit({
                        "request_no": request_no,
                        "estimate_no": "",
                        "result": "NG",
                        "detail": str(error),
                    })

                except Exception as error:
                    failed += 1
                    self.item_finished.emit({
                        "request_no": request_no,
                        "estimate_no": "",
                        "result": "失敗",
                        "detail": str(error),
                    })

            was_cancelled = self.is_cancel_requested()

            if ledger_changed:
                self.progress.emit(
                    "管理台帳をアップロードしています..."
                )
                onedrive.upload_replace(
                    local_file_path=temp_ledger_path,
                    item=ledger_item,
                )

            summary = {
                "total": total,
                "processed": (
                    success + ng + failed
                ),
                "success": success,
                "ng": ng,
                "failed": failed,
                "cancelled": was_cancelled,
            }

            if was_cancelled:
                self.cancelled.emit(summary)
            else:
                self.progress.emit(
                    "TG一括処理が完了しました。"
                )
                self.finished.emit(summary)

        except Exception as error:
            self.failed.emit(str(error))

        finally:
            if (
                temp_ledger_path
                and os.path.exists(temp_ledger_path)
            ):
                try:
                    os.remove(temp_ledger_path)
                except OSError:
                    pass

    def find_excel_files(self):
        return sorted(
            path
            for path in self.input_folder.glob("*.xlsm")
            if not path.name.startswith("~$")
        )

    @staticmethod
    def extract_request_no(pdf_path: Path) -> str:
        return re.sub(
            r"^仕様書_",
            "",
            pdf_path.stem,
            flags=re.IGNORECASE,
        ).strip()

    def find_excel_for_pdf(
        self,
        pdf_path: Path,
        excel_files,
    ):
        key = self.extract_request_no(pdf_path).lower()

        return next(
            (
                path
                for path in excel_files
                if key in path.name.lower()
            ),
            None,
        )

    def create_unique_output_path(
        self,
        base_name: str,
        extension: str,
    ) -> Path:
        candidate = (
            self.output_folder
            / f"{base_name}{extension}"
        )
        number = 2

        while candidate.exists():
            candidate = (
                self.output_folder
                / f"{base_name}_{number}{extension}"
            )
            number += 1

        return candidate
