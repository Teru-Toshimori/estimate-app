import os
import re
from datetime import datetime


from services.excel_automation_helper import ExcelAutomationSession
from services.msr_input_reader import MsrRequest, MsrRequestRow


class MsrLedgerWriter:
    """
    MSRの管理台帳（見積・請求・注文書発行管理台帳）へ
    1明細（見積依頼番号1件）分の見積情報を1行記入する。

    - 「三井E&Sシステム技研」シートの、
      部署名・案件名・見積金額が空の最初の行へ記入する。
    - 見積/請求番号（D列）が空欄なら前行の番号＋1を採番。
      既に記入されていればその番号をそのまま使う。
    - 採番した番号を、対応する出力見積書のK1へ
      「No.XXXXXXXX」の形式で転記する。

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

    # データ開始行（3行目がヘッダー）
    DATA_START_ROW = 4

    # 空行探索の上限
    MAX_SEARCH_ROW = 500

    # 発注期間（例：7/1～9/30）
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
        台帳へ1行記入し、出力見積書のK1へ番号を転記する。

        戻り値：
            row          記入した台帳の行番号
            estimate_no  採番した見積/請求番号

        issuer_name:
            OneDriveへサインインしているメールアドレスを
            利用者一覧Excelで照合して取得した利用者名。
        """

        ledger_path = os.path.abspath(ledger_path)
        estimate_path = os.path.abspath(estimate_path)

        if not os.path.exists(ledger_path):
            raise FileNotFoundError(
                "台帳ファイルが見つかりません。\n"
                f"{ledger_path}"
            )

        if not os.path.exists(estimate_path):
            raise FileNotFoundError(
                "出力済みの見積書が見つかりません。"
                "先に転記実行を行ってください。\n"
                f"{estimate_path}"
            )

        excel_session = ExcelAutomationSession()
        app = None
        ledger_book = None
        estimate_book = None

        try:
            app = excel_session.start()

            # =====================================
            # 台帳へ記入
            # =====================================
            ledger_book = app.books.open(
                ledger_path,
                update_links=False,
                read_only=False,
                ignore_read_only_recommended=True,
                notify=False,
                add_to_mru=False,
            )

            sheet = ledger_book.sheets[self.SHEET_NAME]

            target_row = self._find_empty_row(sheet)

            estimate_no = self._resolve_estimate_no(
                sheet,
                target_row,
            )

            # No（空欄なら前行＋1）
            if sheet.range(f"B{target_row}").value is None:

                previous_no = sheet.range(
                    f"B{target_row - 1}"
                ).value

                if isinstance(previous_no, (int, float)):
                    sheet.range(f"B{target_row}").value = (
                        int(previous_no) + 1
                    )

            # 見積/請求番号
            sheet.range(f"D{target_row}").value = (
                estimate_no
            )

            # 部署名
            sheet.range(f"C{target_row}").value = (
                request.department_upper
            )

            # 案件名（工事名称　（見積依頼番号））
            sheet.range(f"G{target_row}").value = (
                f"{row.construction_name}　"
                f"（{row.request_no}）"
            )

            # 見積金額
            sheet.range(f"H{target_row}").value = (
                row.amount
            )

            # 見積月
            sheet.range(f"I{target_row}").value = (
                self._estimate_month_text(
                    row.order_period
                )
            )

            # 発行日（実行日）
            # 文字列ではなくdatetime値を設定し、
            # Excelの表示形式をユーザー定義「'yy/m/d」にする。
            issue_date_cell = sheet.range(
                f"J{target_row}"
            )
            issue_date_cell.value = datetime.today().date()
            issue_date_cell.number_format = "'yy/m/d"

            # 発行者
            sheet.range(f"K{target_row}").value = (
                str(issuer_name or "").strip()
            )

            ledger_book.save(ledger_path)

            # =====================================
            # 出力見積書のK1へ番号を転記
            # =====================================
            estimate_book = app.books.open(
                estimate_path,
                update_links=False,
                read_only=False,
                ignore_read_only_recommended=True,
                notify=False,
                add_to_mru=False,
            )

            estimate_sheet = estimate_book.sheets[0]

            estimate_sheet.range("K1").value = (
                f"No.{estimate_no}"
            )

            estimate_book.save(estimate_path)

            return {
                "row": target_row,
                "estimate_no": estimate_no,
            }

        finally:

            try:
                if ledger_book:
                    ledger_book.close()
            except Exception:
                pass

            try:
                if estimate_book:
                    estimate_book.close()
            except Exception:
                pass

            excel_session.close()

    # =====================================
    # 空行の探索
    # =====================================
    def _find_empty_row(self, sheet) -> int:
        """
        部署名（C）・案件名（G）・見積金額（H）が
        すべて空の最初の行を返す。
        """

        for row in range(
            self.DATA_START_ROW,
            self.MAX_SEARCH_ROW + 1,
        ):
            values = sheet.range(
                f"C{row}"
            ).value, sheet.range(
                f"G{row}"
            ).value, sheet.range(
                f"H{row}"
            ).value

            if all(v is None for v in values):
                return row

        raise ValueError(
            "台帳に空行が見つかりません。"
        )

    # =====================================
    # 見積/請求番号の決定
    # =====================================
    def _resolve_estimate_no(
        self,
        sheet,
        target_row: int,
    ) -> int:

        current = sheet.range(f"D{target_row}").value

        # 既に番号が入っていればそれを使う
        if isinstance(current, (int, float)):
            return int(current)

        # 空欄なら前行の番号＋1
        previous = sheet.range(
            f"D{target_row - 1}"
        ).value

        if not isinstance(previous, (int, float)):
            raise ValueError(
                "前行に見積/請求番号がないため"
                "採番できません。"
                f"（{target_row - 1}行目）"
            )

        return int(previous) + 1

    # =====================================
    # 見積月の文字列（例：7-9月）
    # =====================================
    def _estimate_month_text(self, period: str) -> str:

        match = self.PERIOD_PATTERN.search(
            period or ""
        )

        if not match:
            return ""

        start_month = int(match.group(1))
        end_month = int(match.group(3))

        if start_month == end_month:
            return f"{start_month}月"

        return f"{start_month}-{end_month}月"
