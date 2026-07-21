import calendar
import os
import re
import shutil
from datetime import datetime

import openpyxl

from services.excel_automation_helper import ExcelAutomationSession
from services.msr_input_reader import MsrRequest, MsrRequestRow


class MsrEstimateWriter:
    """
    MSR見積書を出力する。

    フォーマットファイルは直接変更せず、
    出力先へコピーした上でコピー側の
    以下のセルのみ書き換える。

    出力の単位は「Input明細1行（見積依頼番号1件）＝
    見積書1ファイル」。1つのInputファイルに複数明細が
    あれば、その数だけ出力する。

    K3   見積書作成日（実行日）
    A7   依頼部署①
    F16  作業件名③
    A20  No.（1固定）
    B20  構成内訳③（工事名称＋人月単価）
    H20  数量（暫定固定値。DEFAULT_QUANTITY参照）
    I20  単位（月）
    J20  単価（金額⑤÷数量）
    K20  金額⑤
    E34  見積依頼番号②
    J33  納期/受渡 期日（発注期間の各月末）
    J34  検査完了期日（発注期間の各月末）
    B36  納期（作業期間）④

    出力ファイルには選択した担当者シートのみを残し、
    他の担当者シートは削除する。

    PDF出力はここでは行わない。K1（見積書発行番号）が
    台帳記入前は未確定のため、台帳記入後
    （MsrLedgerWriter側）でPDFを出力する。
    """

    # 構成内訳のデータ行範囲
    DATA_START_ROW = 20
    DATA_END_ROW = 29

    # 数量（決め打ち。根拠不明のため暫定固定値とし、
    # 提出時に誤りと分かれば要修正）
    DEFAULT_QUANTITY = 3

    # 発注期間（例：7/1～9/30）
    PERIOD_PATTERN = re.compile(
        r"(\d{1,2})\s*/\s*(\d{1,2})"
        r"\s*[～〜~\-]\s*"
        r"(\d{1,2})\s*/\s*(\d{1,2})"
    )

    # =====================================
    # 担当者シート（フォーマットのシート名）一覧
    # =====================================
    @staticmethod
    def list_staff_sheets(format_path: str) -> list:
        """
        フォーマットファイルのシート名一覧を取得する。

        担当者名はフォーマットファイル側で変わり得るため、
        コード側に決め打ちせず、都度ここで読み取る。
        """

        format_path = os.path.abspath(format_path)

        workbook = openpyxl.load_workbook(
            format_path,
            read_only=True,
        )

        try:
            return list(workbook.sheetnames)

        finally:
            workbook.close()

    def write(
        self,
        format_path: str,
        output_path: str,
        request: MsrRequest,
        row: MsrRequestRow,
        staff_sheet: str,
    ):

        format_path = os.path.abspath(format_path)
        output_path = os.path.abspath(output_path)

        if not os.path.exists(format_path):
            raise FileNotFoundError(
                "フォーマットファイルが見つかりません。\n"
                f"{format_path}"
            )

        if os.path.exists(output_path):
            os.remove(output_path)

        # フォーマットは改ざんせず、コピーへ書き込む
        shutil.copyfile(format_path, output_path)

        excel_session = ExcelAutomationSession()
        app = None
        book = None

        try:
            app = excel_session.start()

            book = app.books.open(
                output_path,
                update_links=False,
                read_only=False,
                ignore_read_only_recommended=True,
                notify=False,
                add_to_mru=False,
            )

            sheet_names = [
                s.name for s in book.sheets
            ]

            if staff_sheet not in sheet_names:
                raise ValueError(
                    "指定された担当者シートが"
                    "フォーマットファイルに"
                    "見つかりません。\n"
                    f"指定：{staff_sheet}\n"
                    f"フォーマット内のシート："
                    f"{', '.join(sheet_names)}"
                )

            # 担当者シート（案件ごとに選択）
            sheet = book.sheets[staff_sheet]

            # =====================================
            # 見積書作成日（実行日）
            # =====================================
            sheet.range("K3").value = (
                datetime.today().strftime("%Y/%m/%d")
            )

            # =====================================
            # 依頼部署①
            # =====================================
            sheet.range("A7").value = self._add_onchu(
                request.department
            )

            # =====================================
            # 作業件名③
            # =====================================
            sheet.range("F16").value = (
                row.construction_name
            )

            # =====================================
            # 構成内訳③・金額⑤
            # =====================================
            quantity = self.DEFAULT_QUANTITY

            unit_price = row.amount / quantity

            excel_row = self.DATA_START_ROW

            # No.（1固定）
            sheet.range(f"A{excel_row}").value = 1

            sheet.range(f"B{excel_row}").value = (
                f"{row.construction_name}　"
                f"人月単価 "
                f"{self._format_man(unit_price)}万円"
            )

            # 数量（暫定固定値）・単位・単価
            sheet.range(f"H{excel_row}").value = quantity

            sheet.range(f"I{excel_row}").value = "月"

            sheet.range(f"J{excel_row}").value = unit_price

            sheet.range(f"K{excel_row}").value = row.amount

            # =====================================
            # 見積依頼番号②
            # =====================================
            sheet.range("E34").value = row.request_no

            # =====================================
            # 納期/受渡・検査完了期日
            # （発注期間の各月末）
            # =====================================
            month_ends = self._month_end_dates(
                row.order_period,
                request.request_year,
            )

            if month_ends:
                sheet.range("J33").value = month_ends
                sheet.range("J34").value = month_ends

            # =====================================
            # 納期（作業期間）④
            # =====================================
            period_text = self._period_text(
                row.order_period,
                request.request_year,
            )

            if period_text:
                sheet.range("B36").value = (
                    f"納期（作業期間） {period_text}"
                )

            # =====================================
            # 担当者シート以外は削除
            # =====================================
            for other_sheet in list(book.sheets):

                if other_sheet.name != staff_sheet:
                    other_sheet.delete()

            book.save(output_path)

            # PDF出力は行わない。
            # 台帳記入でK1（見積書発行番号）を
            # 反映した後（MsrLedgerWriter側）で
            # 出力する。

        finally:

            try:
                if book:
                    book.close()
            except Exception:
                pass

            excel_session.close()

    # =====================================
    # 御中追加
    # =====================================
    def _add_onchu(self, department):

        department = (department or "").strip()

        if not department:
            return ""

        if department.endswith("御中"):
            return "　" + department

        return "　" + department + "　御中"

    # =====================================
    # 発注期間の解析
    # =====================================
    def _parse_period(self, period: str):

        match = self.PERIOD_PATTERN.search(
            period or ""
        )

        if not match:
            return None

        start_month, start_day, end_month, end_day = (
            int(x) for x in match.groups()
        )

        return start_month, start_day, end_month, end_day

    # =====================================
    # 月数の算出
    # =====================================
    def _count_months(self, period: str) -> int:

        parsed = self._parse_period(period)

        if not parsed:
            return 0

        start_month, _, end_month, _ = parsed

        months = end_month - start_month + 1

        # 年をまたぐ場合（例：11/1～1/31）
        if months <= 0:
            months += 12

        return months

    # =====================================
    # 各月末日の文字列
    # =====================================
    def _month_end_dates(
        self,
        period: str,
        year: int,
    ) -> str:

        parsed = self._parse_period(period)

        if not parsed:
            return ""

        start_month, _, _, _ = parsed

        months = self._count_months(period)

        dates = []

        for i in range(months):

            month_index = start_month + i

            target_year = (
                year + (month_index - 1) // 12
            )

            target_month = (month_index - 1) % 12 + 1

            last_day = calendar.monthrange(
                target_year,
                target_month,
            )[1]

            dates.append(
                f"{target_year}/{target_month}/{last_day}"
            )

        return "、".join(dates)

    # =====================================
    # 作業期間の文字列
    # =====================================
    def _period_text(
        self,
        period: str,
        year: int,
    ) -> str:

        parsed = self._parse_period(period)

        if not parsed:
            return ""

        start_month, start_day, end_month, end_day = (
            parsed
        )

        return (
            f"{year}年{start_month}月{start_day}日"
            f"～{end_month}月{end_day}日"
        )

    # =====================================
    # 万円表記
    # =====================================
    def _format_man(self, value: float) -> str:

        man = value / 10000

        if abs(man - round(man)) < 1e-9:
            return str(int(round(man)))

        return f"{man:g}"
