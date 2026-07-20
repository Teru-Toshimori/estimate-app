import os
import threading

from PySide6.QtCore import QObject, Signal, Slot

from services.msr_ledger_writer import MsrLedgerWriter
from services.onedrive_service import OneDriveService
from services.user_master_reader import UserMasterReader


class MsrLedgerUpdateWorker(QObject):
    """
    OneDrive／SharePoint上のMSR管理台帳へ、
    複数の見積書情報をまとめて記入するWorker。

    処理の開始時に、Microsoft Graphから現在の利用者情報を取得し、
    利用者一覧Excelをメールアドレスで検索する。
    取得した利用者名を、管理台帳K列の発行者へ記入する。
    """

    finished = Signal(list)
    cancelled = Signal(list)
    failed = Signal(str)
    progress = Signal(str)

    def __init__(
        self,
        share_url,
        user_master_url,
        jobs,
        device_flow_callback=None,
    ):
        super().__init__()

        self.share_url = str(
            share_url or ""
        ).strip()

        self.user_master_url = str(
            user_master_url or ""
        ).strip()

        self.jobs = jobs
        self.device_flow_callback = (
            device_flow_callback
        )

        self._cancel_event = threading.Event()

    # =====================================
    # 中止
    # =====================================
    def request_cancel(self) -> None:
        """
        現在の安全な処理単位が終了した時点で停止する。
        """

        self._cancel_event.set()

    def is_cancel_requested(self) -> bool:

        return self._cancel_event.is_set()

    # =====================================
    # 一括処理
    # =====================================
    @Slot()
    def run(self) -> None:

        ledger_temp_path = None

        try:
            if self.is_cancel_requested():
                self.cancelled.emit([])
                return

            self.progress.emit(
                "Microsoftアカウントを確認しています..."
            )

            onedrive_service = OneDriveService(
                device_flow_callback=(
                    self.device_flow_callback
                )
            )

            ledger_writer = MsrLedgerWriter()

            # =====================================
            # 利用者名を取得
            # =====================================
            self.progress.emit(
                "利用者情報を取得しています..."
            )

            issuer_name = self.resolve_issuer_name(
                onedrive_service
            )

            if self.is_cancel_requested():
                self.cancelled.emit([])
                return

            # =====================================
            # 管理台帳をダウンロード
            # =====================================
            self.progress.emit(
                "管理台帳をダウンロードしています..."
            )

            ledger_temp_path, ledger_item = (
                onedrive_service.download_to_temp(
                    self.share_url
                )
            )

            results = []
            changed = False

            # =====================================
            # 管理台帳へ順番に記入
            # =====================================
            for job in self.jobs:

                if self.is_cancel_requested():
                    break

                self.progress.emit(
                    "台帳記入中："
                    f"{job['file_name']}"
                )

                common_result = {
                    "file_name": job["file_name"],
                    "request_no": job["request_no"],
                    "result_key": job["result_key"],
                }

                try:
                    result = ledger_writer.write(
                        ledger_path=ledger_temp_path,
                        estimate_path=(
                            job["estimate_path"]
                        ),
                        request=job["request"],
                        issuer_name=issuer_name,
                    )

                    changed = True

                    results.append({
                        **common_result,
                        "success": True,
                        "row": result["row"],
                        "estimate_no": (
                            result["estimate_no"]
                        ),
                    })

                except Exception as error:
                    results.append({
                        **common_result,
                        "success": False,
                        "error": str(error),
                    })

            was_cancelled = (
                self.is_cancel_requested()
            )

            # =====================================
            # 管理台帳をアップロード
            # =====================================
            if changed:
                if was_cancelled:
                    self.progress.emit(
                        "中止前に完了した内容を"
                        "管理台帳へアップロードしています..."
                    )
                else:
                    self.progress.emit(
                        "管理台帳をアップロードしています..."
                    )

                onedrive_service.upload_replace(
                    local_file_path=(
                        ledger_temp_path
                    ),
                    item=ledger_item,
                )

            if was_cancelled:
                self.progress.emit(
                    "管理台帳の更新を中止しました。"
                )
                self.cancelled.emit(results)
                return

            self.progress.emit(
                "管理台帳の更新が完了しました。"
            )
            self.finished.emit(results)

        except Exception as error:
            self.failed.emit(str(error))

        finally:
            if (
                ledger_temp_path
                and os.path.exists(
                    ledger_temp_path
                )
            ):
                try:
                    os.remove(
                        ledger_temp_path
                    )
                except OSError:
                    pass

    # =====================================
    # 利用者一覧から発行者名を取得
    # =====================================
    def resolve_issuer_name(
        self,
        onedrive_service: OneDriveService,
    ) -> str:
        """
        サインイン中のMicrosoftアカウントから
        メールアドレス候補を取得し、
        利用者一覧Excelを検索して利用者名を返す。
        """

        if not self.user_master_url:
            raise ValueError(
                "利用者一覧URLが入力されていません。"
            )

        if not self.user_master_url.startswith(
            ("https://", "http://")
        ):
            raise ValueError(
                "利用者一覧URLの形式が"
                "正しくありません。\n\n"
                f"{self.user_master_url}"
            )

        user = onedrive_service.get_current_user()

        mail = str(
            user.get("mail") or ""
        ).strip()

        user_principal_name = str(
            user.get("userPrincipalName") or ""
        ).strip()

        other_mails = (
            user.get("otherMails")
            or []
        )

        email_candidates = []

        for email in [
            mail,
            user_principal_name,
            *other_mails,
        ]:
            normalized = str(
                email or ""
            ).strip()

            if (
                normalized
                and normalized.lower()
                not in {
                    value.lower()
                    for value in email_candidates
                }
            ):
                email_candidates.append(
                    normalized
                )

        if not email_candidates:
            raise ValueError(
                "サインイン中のMicrosoftアカウントから"
                "メールアドレスを取得できませんでした。"
            )

        user_master_temp_path = None

        try:
            user_master_temp_path, _ = (
                onedrive_service.download_to_temp(
                    self.user_master_url
                )
            )

            reader = UserMasterReader()

            issuer_name = reader.find_user_name(
                excel_path=(
                    user_master_temp_path
                ),
                email_addresses=(
                    email_candidates
                ),
            )

            issuer_name = str(
                issuer_name or ""
            ).strip()

            if issuer_name:
                return issuer_name

            raise ValueError(
                "利用者一覧Excelに、サインイン中の"
                "メールアドレスと一致する利用者が"
                "見つかりませんでした。\n\n"
                f"mail: {mail or 'なし'}\n"
                "userPrincipalName: "
                f"{user_principal_name or 'なし'}"
            )

        finally:
            if (
                user_master_temp_path
                and os.path.exists(
                    user_master_temp_path
                )
            ):
                try:
                    os.remove(
                        user_master_temp_path
                    )
                except OSError:
                    pass
