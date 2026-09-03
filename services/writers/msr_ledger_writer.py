import os
import re
from datetime import datetime

from services.excel.excel_automation_helper import (
    ExcelAutomationSession,
)
from services.readers.msr_input_reader import (
    MsrRequest,
    MsrRequestRow,
)


class MsrLedgerWriter:
    """
    MSRの管理台帳（見積・請求・注文書発行管理台帳）へ
    1明細（見積依頼番号1件）分の見積情報を1行記入する。

    - 「三井E&Sシステム技研」シートの、
      部署名・案件名・見積金額が空の最初の行へ記入する。
    - 見積/請求番号（D列）が空欄なら前行の番号＋1を採番。
      既に記入されていれば、その番号をそのまま使う。
    - 採番した番号を、対応する出力見積書のK1へ
      「No.XXXXXXXX」の形式で転記する。
    - K1確定後の最終状態で、同名のPDFも出力する。

    記入する列：
    B  No（空欄なら前行＋1）
    C  部署名（Inputの依頼元 上段）
    D  見積/請求番号
    G  案件名（工事名称　（見積依頼番号））
    H  見積金額
    I  見積月（発注期間から「7-9月」形式）
    J  発行日（実行日）
    K  発行者（利用者一覧から取得した利用者名）
    """

    SHEET_NAME = "三井E&Sシステム技研"

    DATA_START_ROW = 4
    MAX_SEARCH_ROW = 500

    PERIOD_PATTERN = re.compile(
        r"(\d{1,2})\s*/\s*(\d{1,2})"
        r"\s*[～〜~\-]\s*"
        r"(\d{1,2})\s*/\s*(\d{1,2})"
    )

    def write(
        self,
        ledger_path: str,
        estimate_path: str,
        request: MsrRequest,
        row: MsrRequestRow,
        issuer_name: str,
    ) -> dict:
        """
        台帳へ1行記入し、
        出力見積書のK1へ番号を転記する。

        戻り値：
            {
                "row": 記入した台帳の行番号,
                "estimate_no": 採番した見積/請求番号,
            }
        """

        ledger_path = os.path.abspath(
            ledger_path
        )

        estimate_path = os.path.abspath(
            estimate_path
        )

        if not os.path.exists(
            ledger_path
        ):
            raise FileNotFoundError(
                "台帳ファイルが見つかりません。\n"
                f"{ledger_path}"
            )

        if not os.path.exists(
            estimate_path
        ):
            raise FileNotFoundError(
                "出力済みの見積書が見つかりません。"
                "先に転記実行を行ってください。\n"
                f"{estimate_path}"
            )

        excel_session = (
            ExcelAutomationSession()
        )

        app = None
        ledger_book = None
        estimate_book = None

        try:
            app = (
                excel_session.start()
            )

            # =====================================
            # 台帳を開く
            # =====================================
            try:
                ledger_book = (
                    app.books.open(
                        ledger_path,
                        update_links=False,
                        read_only=False,
                        ignore_read_only_recommended=True,
                        notify=False,
                        add_to_mru=False,
                    )
                )

            except Exception as error:
                raise RuntimeError(
                    self._build_ledger_open_error_message(
                        error
                    )
                ) from error

            sheet = (
                ledger_book.sheets[
                    self.SHEET_NAME
                ]
            )

            target_row = (
                self._find_empty_row(
                    sheet
                )
            )

            estimate_no = (
                self._resolve_estimate_no(
                    sheet,
                    target_row,
                )
            )

            # =====================================
            # B列：No
            # =====================================
            current_no = (
                sheet.range(
                    f"B{target_row}"
                ).value
            )

            if current_no is None:

                previous_no = (
                    sheet.range(
                        f"B{target_row - 1}"
                    ).value
                )

                if isinstance(
                    previous_no,
                    (int, float),
                ):
                    sheet.range(
                        f"B{target_row}"
                    ).value = (
                        int(previous_no)
                        + 1
                    )

            # =====================================
            # D列：見積/請求番号
            # =====================================
            sheet.range(
                f"D{target_row}"
            ).value = estimate_no

            # =====================================
            # C列：部署名
            # =====================================
            sheet.range(
                f"C{target_row}"
            ).value = (
                request.department_upper
            )

            # =====================================
            # G列：案件名
            # =====================================
            sheet.range(
                f"G{target_row}"
            ).value = (
                f"{row.construction_name}　"
                f"（{row.request_no}）"
            )

            # =====================================
            # H列：見積金額
            # =====================================
            sheet.range(
                f"H{target_row}"
            ).value = row.amount

            # =====================================
            # I列：見積月
            # =====================================
            sheet.range(
                f"I{target_row}"
            ).value = (
                self._estimate_month_text(
                    row.order_period
                )
            )

            # =====================================
            # J列：発行日
            # =====================================
            issue_date_cell = (
                sheet.range(
                    f"J{target_row}"
                )
            )

            issue_date_cell.value = (
                datetime.today().date()
            )

            issue_date_cell.number_format = (
                "'yy/m/d"
            )

            # =====================================
            # K列：発行者
            # =====================================
            sheet.range(
                f"K{target_row}"
            ).value = (
                str(
                    issuer_name
                    or ""
                ).strip()
            )

            # =====================================
            # 台帳保存
            # =====================================
            ledger_book.save(
                ledger_path
            )

            # =====================================
            # 出力見積書を開く
            # =====================================
            estimate_book = (
                app.books.open(
                    estimate_path,
                    update_links=False,
                    read_only=False,
                    ignore_read_only_recommended=True,
                    notify=False,
                    add_to_mru=False,
                )
            )

            estimate_sheet = (
                estimate_book.sheets[0]
            )

            estimate_sheet.range(
                "K1"
            ).value = (
                f"No.{estimate_no}"
            )

            estimate_book.save(
                estimate_path
            )

            # =====================================
            # PDF出力
            # =====================================
            pdf_path = (
                os.path.splitext(
                    estimate_path
                )[0]
                + ".pdf"
            )

            estimate_sheet.api.ExportAsFixedFormat(
                Type=0,
                Filename=pdf_path,
            )

            return {
                "row": target_row,
                "estimate_no": estimate_no,
            }

        finally:
            # =====================================
            # 出力見積書を閉じる
            # =====================================
            try:
                if (
                    estimate_book
                    is not None
                ):
                    estimate_book.close()

            except Exception:
                pass

            # =====================================
            # 台帳を閉じる
            # =====================================
            try:
                if (
                    ledger_book
                    is not None
                ):
                    ledger_book.close()

            except Exception:
                pass

            # =====================================
            # Excel終了
            # =====================================
            excel_session.close()

    # =====================================
    # MSR管理台帳OPEN失敗メッセージ
    # =====================================
    def _build_ledger_open_error_message(
        self,
        error: Exception,
    ) -> str:
        """
        MSR管理台帳をExcelで開けなかった場合に、
        利用者が確認すべき内容を分かりやすく返す。

        過去に、台帳内部へ古い外部リンク情報が残り、
        externalLink XMLの一部に不整合が発生したことで
        ExcelのWorkbooks.Openが失敗した事例があった。

        ただし、Open失敗の原因は外部リンクだけとは限らないため、
        「可能性があります」として案内する。
        """

        error_text = str(
            error
        ).strip()

        message = (
            "MSR管理台帳を開くことができませんでした。\n\n"
            "管理台帳に外部ファイルを参照する「外部リンク」情報が残っており、\n"
            "そのリンク情報の一部にExcelが不整合を検出している可能性があります。\n\n"
            "【確認してください】\n"
            "1. OneDrive上のMSR管理台帳をExcelで直接開いてください。\n"
            "2. 「ファイルに問題が見つかりました」など、"
            "修復を求めるメッセージが表示された場合は、"
            "修復を実行してください。\n"
            "3. 修復後、そのExcelファイルを上書き保存してください。\n"
            "4. 保存後、見積書作成ツールでもう一度実行してください。\n\n"
            "※ 外部リンクが存在するだけでは問題ありません。\n"
            "   Excelが外部リンク情報の不整合を検出した場合に"
            "発生する可能性があります。\n"
            "※ ファイルがExcelで開かれている、アクセス権がない、"
            "Excel側で問題が発生している場合などでも、"
            "同様に台帳を開けないことがあります。\n\n"
            "対象：MSR管理台帳"
        )

        if error_text:
            message += (
                "\n\n"
                "エラー詳細：\n"
                f"{error_text}"
            )

        return message

    # =====================================
    # 空行探索
    # =====================================
    def _find_empty_row(
        self,
        sheet,
    ) -> int:
        """
        C・G・H列がすべて空の
        最初の行を返す。
        """

        for row_number in range(
            self.DATA_START_ROW,
            self.MAX_SEARCH_ROW + 1,
        ):

            department = (
                sheet.range(
                    f"C{row_number}"
                ).value
            )

            project_name = (
                sheet.range(
                    f"G{row_number}"
                ).value
            )

            amount = (
                sheet.range(
                    f"H{row_number}"
                ).value
            )

            values = (
                department,
                project_name,
                amount,
            )

            if all(
                value is None
                for value in values
            ):
                return row_number

        raise ValueError(
            "台帳に空行が見つかりません。"
        )

    # =====================================
    # 見積/請求番号決定
    # =====================================
    def _resolve_estimate_no(
        self,
        sheet,
        target_row: int,
    ) -> int:
        """
        対象行D列に番号があれば使用。

        空欄なら前行の番号＋1。
        """

        current = (
            sheet.range(
                f"D{target_row}"
            ).value
        )

        if isinstance(
            current,
            (int, float),
        ):
            return int(
                current
            )

        previous = (
            sheet.range(
                f"D{target_row - 1}"
            ).value
        )

        if not isinstance(
            previous,
            (int, float),
        ):
            raise ValueError(
                "前行に見積/請求番号がないため"
                "採番できません。"
                f"（{target_row - 1}行目）"
            )

        return (
            int(previous)
            + 1
        )

    # =====================================
    # 見積月変換
    # =====================================
    def _estimate_month_text(
        self,
        period: str,
    ) -> str:
        """
        例：

        7/1～9/30
        ↓
        7-9月

        7/1～7/31
        ↓
        7月
        """

        match = (
            self.PERIOD_PATTERN.search(
                period
                or ""
            )
        )

        if not match:
            return ""

        start_month = int(
            match.group(1)
        )

        end_month = int(
            match.group(3)
        )

        if (
            start_month
            == end_month
        ):
            return (
                f"{start_month}月"
            )

        return (
            f"{start_month}-"
            f"{end_month}月"
        )