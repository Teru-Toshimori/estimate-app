import os
import re
import shutil
import tempfile
from pathlib import Path

from PySide6.QtCore import QObject, Signal, Slot

from services.template_resolver import TemplateResolver
from services.excel_writer import ExcelWriter
from services.ledger_writer import (
    DuplicateApplicationNoError,
    LedgerWriter,
)
from services.onedrive_service import OneDriveService
from services.pdf_reader import PDFReader


class BatchEstimateWorker(QObject):
    """
    特調TBの見積書を一括作成するバックグラウンド処理。

    処理概要:
    1. PDFファイルの事前チェック
    2. OneDrive／SharePointから管理台帳を1回ダウンロード
    3. 利用者一覧から発行者名を取得
    4. PDFを1件ずつ解析
    5. 業務計画書Noの重複確認
    6. 重複していなければ管理台帳へ採番・追記
    7. 見積書Excel・PDFを出力
    8. 全件終了後、管理台帳を1回アップロード
    9. 管理台帳更新成功後に結果を画面へ通知

    業務計画書Noが登録済みの場合:
    ・採番しない
    ・管理台帳へ追記しない
    ・Excel／PDFを作成しない
    ・リザルト画面へNGを表示する
    """

    # PDF1件ごとの処理結果
    item_finished = Signal(dict)

    # 全件処理完了
    batch_finished = Signal(dict)

    # 一括処理全体の失敗
    batch_failed = Signal(str)

    # 状態メッセージ
    progress = Signal(str)

    # 件数ベースの進捗
    progress_count = Signal(
        int,
        int,
        str,
    )

    # Microsoft認証画面の表示要求
    device_login_requested = Signal(dict)

    def __init__(
        self,
        pdf_files: list[Path],
        share_url: str,
        user_master_url: str,
        output_folder: str,
    ):
        super().__init__()

        self.pdf_files = [
            Path(path)
            for path in pdf_files
        ]

        self.share_url = share_url
        self.user_master_url = user_master_url

        self.output_folder = Path(
            output_folder
        )

        self.is_cancelled = False

    # =====================================
    # 一括処理
    # =====================================
    @Slot()
    def run(self):

        total_count = len(
            self.pdf_files
        )

        success_count = 0
        failed_count = 0
        ng_count = 0

        ledger_temp_path = None
        ledger_item = None
        ledger_backup_path = None

        # 台帳アップロード後に通知する成功結果
        pending_results: list[dict] = []

        # 通常エラーの結果
        failed_results: list[dict] = []

        # 重複などのNG結果
        ng_results: list[dict] = []

        # 今回作成した出力ファイル
        created_output_files: list[Path] = []

        # 既存出力ファイルのバックアップ
        output_backups: dict[
            Path,
            Path,
        ] = {}

        try:
            # =====================================
            # 進捗初期化
            # =====================================
            self.progress_count.emit(
                0,
                total_count,
                "",
            )

            # =====================================
            # PDF事前チェック
            # =====================================
            self.progress.emit(
                "PDFファイルを確認しています..."
            )

            self.validate_pdf_files()

            # =====================================
            # 出力フォルダ準備
            # =====================================
            self.output_folder.mkdir(
                parents=True,
                exist_ok=True,
            )

            # =====================================
            # 見積書テンプレート取得
            # =====================================
            template_path = TemplateResolver.resolve("tb")

            # =====================================
            # サービス生成
            # =====================================
            pdf_reader = PDFReader()
            ledger_writer = LedgerWriter()
            excel_writer = ExcelWriter()

            onedrive_service = OneDriveService(
                device_flow_callback=(
                    self.request_device_login
                ),
                user_master_url=(
                    self.user_master_url
                ),
            )

            # =====================================
            # 管理台帳を1回だけダウンロード
            # =====================================
            self.progress.emit(
                "管理台帳をダウンロードしています..."
            )

            ledger_temp_path, ledger_item = (
                onedrive_service.download_to_temp(
                    self.share_url
                )
            )

            # =====================================
            # 発行者名を1回だけ取得
            # =====================================
            self.progress.emit(
                "発行者情報を取得しています..."
            )

            issuer_name = (
                onedrive_service
                .get_issuer_name_from_master()
            )

            # =====================================
            # PDFを順番に処理
            # =====================================
            for index, pdf_path in enumerate(
                self.pdf_files,
                start=1,
            ):
                if self.is_cancelled:
                    break

                voucher_no = (
                    pdf_path.stem.strip()
                )

                estimate_no = ""

                # 出力ファイル名
                # 既存ファイルは残し、同名の場合は
                # 「_2」「_3」…を付けて新規保存する。
                (
                    output_excel_path,
                    output_pdf_path,
                ) = self.create_unique_output_paths(
                    file_name_base=(
                        f"「見積書」{voucher_no}"
                    )
                )

                self.progress.emit(
                    f"{index}/{total_count}件目を"
                    "処理しています："
                    f"{pdf_path.name}"
                )

                self.progress_count.emit(
                    index - 1,
                    total_count,
                    pdf_path.name,
                )

                # =====================================
                # PDF処理前の台帳をバックアップ
                # =====================================
                ledger_backup_path = (
                    f"{ledger_temp_path}.backup"
                )

                shutil.copy2(
                    ledger_temp_path,
                    ledger_backup_path,
                )

                try:
                    # =====================================
                    # PDF解析
                    # =====================================
                    data = pdf_reader.parse(
                        str(pdf_path)
                    )

                    data.voucher_no = (
                        voucher_no
                    )

                    # =====================================
                    # 重複確認・管理台帳採番
                    # =====================================
                    ledger_result = (
                        ledger_writer.write(
                            excel_path=(
                                ledger_temp_path
                            ),
                            data=data,
                            issuer_name=(
                                issuer_name
                            ),
                        )
                    )

                    estimate_no = str(
                        ledger_result.get(
                            "estimate_no",
                            "",
                        )
                    )

                    issue_date = str(
                        ledger_result.get(
                            "issue_date",
                            "",
                        )
                    )

                    data.estimate_no = (
                        estimate_no
                    )

                    data.issue_date = (
                        issue_date
                    )

                    # =====================================
                    # 既存ファイルは残す
                    # =====================================
                    # 出力パス作成時に連番を付けているため、
                    # 既存ファイルの退避・上書きは行わない。

                    # =====================================
                    # Excel・PDF出力
                    # =====================================
                    excel_writer.write(
                        str(template_path),
                        str(output_excel_path),
                        data,
                    )

                    # =====================================
                    # 出力確認
                    # =====================================
                    if not output_excel_path.exists():
                        raise RuntimeError(
                            "見積書Excelが"
                            "作成されませんでした。\n"
                            f"{output_excel_path}"
                        )

                    if not output_pdf_path.exists():
                        raise RuntimeError(
                            "見積書PDFが"
                            "作成されませんでした。\n"
                            f"{output_pdf_path}"
                        )

                    success_count += 1

                    pending_results.append(
                        {
                            "voucher_no": (
                                voucher_no
                            ),
                            "estimate_no": (
                                estimate_no
                            ),
                            "result": "成功",
                            "message": "",
                            "excel_path": str(
                                output_excel_path
                            ),
                            "pdf_path": str(
                                output_pdf_path
                            ),
                        }
                    )

                    created_output_files.extend(
                        [
                            output_excel_path,
                            output_pdf_path,
                        ]
                    )

                    self.remove_file_safely(
                        ledger_backup_path
                    )

                    ledger_backup_path = None

                # =====================================
                # 業務計画書No重複
                # =====================================
                except DuplicateApplicationNoError as error:
                    ng_count += 1

                    # 台帳を処理前状態へ戻す
                    if (
                        ledger_backup_path
                        and os.path.exists(
                            ledger_backup_path
                        )
                    ):
                        shutil.copy2(
                            ledger_backup_path,
                            ledger_temp_path,
                        )

                    # 念のため出力途中ファイルを削除
                    self.remove_file_safely(
                        output_excel_path
                    )

                    self.remove_file_safely(
                        output_pdf_path
                    )

                    # 既存ファイルが退避されていれば復元
                    self.restore_output_backup(
                        output_excel_path,
                        output_backups,
                    )

                    self.restore_output_backup(
                        output_pdf_path,
                        output_backups,
                    )

                    ng_results.append(
                        {
                            "voucher_no": (
                                voucher_no
                            ),
                            "estimate_no": "",
                            "result": "NG",
                            "message": str(error),
                            "excel_path": "",
                            "pdf_path": "",
                        }
                    )

                    self.remove_file_safely(
                        ledger_backup_path
                    )

                    ledger_backup_path = None

                # =====================================
                # その他のエラー
                # =====================================
                except Exception as error:
                    failed_count += 1

                    # 台帳を処理前状態へ戻す
                    if (
                        ledger_backup_path
                        and os.path.exists(
                            ledger_backup_path
                        )
                    ):
                        shutil.copy2(
                            ledger_backup_path,
                            ledger_temp_path,
                        )

                    # 出力途中ファイルを削除
                    self.remove_file_safely(
                        output_excel_path
                    )

                    self.remove_file_safely(
                        output_pdf_path
                    )

                    # 既存ファイルを復元
                    self.restore_output_backup(
                        output_excel_path,
                        output_backups,
                    )

                    self.restore_output_backup(
                        output_pdf_path,
                        output_backups,
                    )

                    failed_results.append(
                        {
                            "voucher_no": (
                                voucher_no
                            ),
                            "estimate_no": (
                                estimate_no
                            ),
                            "result": "失敗",
                            "message": str(error),
                            "excel_path": "",
                            "pdf_path": "",
                        }
                    )

                    self.remove_file_safely(
                        ledger_backup_path
                    )

                    ledger_backup_path = None

                # 成功・NG・失敗にかかわらず
                # このPDFの処理完了を通知
                self.progress_count.emit(
                    index,
                    total_count,
                    pdf_path.name,
                )

            processed_count = (
                success_count
                + failed_count
                + ng_count
            )

            # =====================================
            # 成功案件がある場合だけ台帳アップロード
            # =====================================
            if success_count > 0:
                self.progress.emit(
                    "更新した管理台帳を"
                    "OneDrive／SharePointへ"
                    "アップロードしています..."
                )

                onedrive_service.upload_replace(
                    local_file_path=(
                        ledger_temp_path
                    ),
                    item=ledger_item,
                )

            # =====================================
            # 台帳更新成功後、古い出力を削除
            # =====================================
            self.delete_output_backups(
                output_backups
            )

            # =====================================
            # 結果を画面へ通知
            # =====================================
            for result in pending_results:
                self.item_finished.emit(
                    result
                )

            for result in ng_results:
                self.item_finished.emit(
                    result
                )

            for result in failed_results:
                self.item_finished.emit(
                    result
                )

            # =====================================
            # 完了通知
            # =====================================
            if self.is_cancelled:
                self.progress.emit(
                    "一括処理を中止しました。"
                )
            else:
                self.progress.emit(
                    "一括処理が完了しました。"
                )

            self.batch_finished.emit(
                {
                    "total": total_count,
                    "processed": (
                        processed_count
                    ),
                    "success": success_count,
                    "failed": failed_count,
                    "ng": ng_count,
                    "cancelled": (
                        self.is_cancelled
                    ),
                }
            )

        except Exception as error:
            # =====================================
            # 今回作成した見積書を削除
            # =====================================
            for output_path in (
                created_output_files
            ):
                self.remove_file_safely(
                    output_path
                )

            # =====================================
            # 処理前の既存見積書を復元
            # =====================================
            restore_errors = (
                self.restore_all_output_backups(
                    output_backups
                )
            )

            error_message = (
                "一括処理を完了できなかったため、"
                "今回作成した見積書は"
                "削除しました。\n"
                "処理前から存在していた見積書は"
                "復元しました。\n\n"
                f"{error}"
            )

            if restore_errors:
                error_message += (
                    "\n\n一部の既存見積書を"
                    "復元できませんでした。\n"
                    + "\n".join(
                        restore_errors
                    )
                )

            self.batch_failed.emit(
                error_message
            )

        finally:
            self.remove_file_safely(
                ledger_backup_path
            )

            self.remove_file_safely(
                ledger_temp_path
            )

            self.cleanup_backup_directories(
                output_backups
            )

    # =====================================
    # PDF事前チェック
    # =====================================
    def validate_pdf_files(
        self,
    ) -> None:
        """
        一括処理開始前にPDF一覧を検査する。
        """

        if not self.pdf_files:
            raise ValueError(
                "処理対象のPDFがありません。"
            )

        errors: list[str] = []
        voucher_numbers: set[str] = set()

        invalid_name_pattern = re.compile(
            r'[<>:"/\\|?*]'
        )

        reserved_names = {
            "CON",
            "PRN",
            "AUX",
            "NUL",
            *{
                f"COM{number}"
                for number in range(
                    1,
                    10,
                )
            },
            *{
                f"LPT{number}"
                for number in range(
                    1,
                    10,
                )
            },
        }

        for pdf_path in self.pdf_files:
            path = Path(
                pdf_path
            )

            if not path.exists():
                errors.append(
                    "ファイルが見つかりません："
                    f"{path}"
                )
                continue

            if not path.is_file():
                errors.append(
                    "ファイルではありません："
                    f"{path}"
                )
                continue

            if path.suffix.lower() != ".pdf":
                errors.append(
                    "PDFではありません："
                    f"{path.name}"
                )
                continue

            try:
                file_size = (
                    path.stat().st_size
                )

            except OSError as error:
                errors.append(
                    "ファイル情報を取得できません："
                    f"{path.name}（{error}）"
                )
                continue

            if file_size <= 0:
                errors.append(
                    f"空のPDFです：{path.name}"
                )

            voucher_no = (
                path.stem.strip()
            )

            if not voucher_no:
                errors.append(
                    "伝票番号を取得できません："
                    f"{path.name}"
                )
                continue

            if invalid_name_pattern.search(
                voucher_no
            ):
                errors.append(
                    "伝票番号にファイル名として"
                    "使用できない文字があります："
                    f"{voucher_no}"
                )
                continue

            if (
                voucher_no.endswith(" ")
                or voucher_no.endswith(".")
            ):
                errors.append(
                    "伝票番号の末尾に使用できない"
                    "空白またはピリオドがあります："
                    f"{voucher_no}"
                )
                continue

            if (
                voucher_no.upper()
                in reserved_names
            ):
                errors.append(
                    "伝票番号がWindowsの予約名です："
                    f"{voucher_no}"
                )
                continue

            normalized_voucher = (
                voucher_no.casefold()
            )

            if (
                normalized_voucher
                in voucher_numbers
            ):
                errors.append(
                    "同じ伝票番号のPDFが"
                    "複数あります："
                    f"{voucher_no}"
                )
            else:
                voucher_numbers.add(
                    normalized_voucher
                )

        if errors:
            raise ValueError(
                "PDFの事前チェックで"
                "問題が見つかりました。\n\n"
                + "\n".join(
                    f"・{error}"
                    for error in errors
                )
            )

    # =====================================
    # 重複しない出力パスを作成
    # =====================================
    def create_unique_output_paths(
        self,
        file_name_base: str,
    ) -> tuple[Path, Path]:
        """
        ExcelまたはPDFのどちらかが既に存在する場合、
        元ファイルを残したまま連番付きの出力パスを返す。

        例:
            「見積書」ITK14642.xlsx
            「見積書」ITK14642_2.xlsx
            「見積書」ITK14642_3.xlsx
        """

        candidate = file_name_base
        number = 2

        while (
            (
                self.output_folder
                / f"{candidate}.xlsx"
            ).exists()
            or (
                self.output_folder
                / f"{candidate}.pdf"
            ).exists()
        ):
            candidate = (
                f"{file_name_base}_{number}"
            )

            number += 1

        return (
            self.output_folder
            / f"{candidate}.xlsx",
            self.output_folder
            / f"{candidate}.pdf",
        )

    # =====================================
    # Microsoft認証要求
    # =====================================
    def request_device_login(
        self,
        flow: dict,
    ):

        self.device_login_requested.emit(
            flow
        )

    # =====================================
    # 処理中止
    # =====================================
    def cancel(self):

        self.is_cancelled = True

    # =====================================
    # 既存出力ファイルを一時退避
    # =====================================
    def backup_existing_output(
        self,
        file_path: Path,
        output_backups: dict[
            Path,
            Path,
        ],
    ) -> None:

        if not file_path.exists():
            return

        if file_path in output_backups:
            return

        try:
            backup_dir = Path(
                tempfile.mkdtemp(
                    prefix=(
                        "estimate_output_backup_"
                    )
                )
            )

            backup_path = (
                backup_dir
                / file_path.name
            )

            shutil.move(
                str(file_path),
                str(backup_path),
            )

            output_backups[
                file_path
            ] = backup_path

        except PermissionError as error:
            raise PermissionError(
                "既存の出力ファイルを"
                "退避できません。\n\n"
                f"{file_path}\n\n"
                "対象ファイルがExcelまたは"
                "PDFビューアーで"
                "開かれていないか"
                "確認してください。"
            ) from error

        except OSError as error:
            raise RuntimeError(
                "既存の出力ファイルを"
                "バックアップできませんでした。\n\n"
                f"{file_path}\n\n"
                f"{error}"
            ) from error

    # =====================================
    # 指定した既存ファイルを復元
    # =====================================
    def restore_output_backup(
        self,
        original_path: Path,
        output_backups: dict[
            Path,
            Path,
        ],
    ) -> None:

        backup_path = output_backups.pop(
            original_path,
            None,
        )

        if backup_path is None:
            return

        try:
            self.remove_file_safely(
                original_path
            )

            original_path.parent.mkdir(
                parents=True,
                exist_ok=True,
            )

            shutil.move(
                str(backup_path),
                str(original_path),
            )

            self.remove_empty_parent(
                backup_path.parent
            )

        except OSError as error:
            raise RuntimeError(
                "既存の見積書を"
                "復元できませんでした。\n\n"
                f"復元先：{original_path}\n"
                f"バックアップ：{backup_path}"
            ) from error

    # =====================================
    # すべての既存ファイルを復元
    # =====================================
    def restore_all_output_backups(
        self,
        output_backups: dict[
            Path,
            Path,
        ],
    ) -> list[str]:

        restore_errors: list[str] = []

        for original_path in list(
            output_backups.keys()
        ):
            try:
                self.restore_output_backup(
                    original_path,
                    output_backups,
                )

            except Exception as error:
                restore_errors.append(
                    f"{original_path}：{error}"
                )

        return restore_errors

    # =====================================
    # 正常終了時に古いバックアップを削除
    # =====================================
    def delete_output_backups(
        self,
        output_backups: dict[
            Path,
            Path,
        ],
    ) -> None:

        for backup_path in list(
            output_backups.values()
        ):
            self.remove_file_safely(
                backup_path
            )

            self.remove_empty_parent(
                backup_path.parent
            )

        output_backups.clear()

    # =====================================
    # バックアップフォルダの後始末
    # =====================================
    def cleanup_backup_directories(
        self,
        output_backups: dict[
            Path,
            Path,
        ],
    ) -> None:

        for backup_path in list(
            output_backups.values()
        ):
            self.remove_empty_parent(
                backup_path.parent
            )

    # =====================================
    # 空フォルダ削除
    # =====================================
    def remove_empty_parent(
        self,
        directory_path: Path,
    ) -> None:

        try:
            if (
                directory_path.exists()
                and directory_path.is_dir()
                and not any(
                    directory_path.iterdir()
                )
            ):
                directory_path.rmdir()

        except OSError:
            pass

    # =====================================
    # ファイルを安全に削除
    # =====================================
    def remove_file_safely(
        self,
        file_path,
    ) -> None:

        if not file_path:
            return

        try:
            path = Path(
                file_path
            )

            if path.exists():
                path.unlink()

        except OSError:
            pass