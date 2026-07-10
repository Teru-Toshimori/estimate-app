import os
import re
from datetime import datetime

import xlwings as xw


class ExcelWriter:

    def write(self, template_path: str, output_path: str, data):

        template_path = os.path.abspath(template_path)
        output_path = os.path.abspath(output_path)

        if os.path.exists(output_path):
            os.remove(output_path)

        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False
        app.screen_updating = False

        template_book = None
        output_book = None

        try:

            # テンプレートを開く
            template_book = app.books.open(template_path)

            # 「フォーマット」シートをコピー
            template_sheet = template_book.sheets["フォーマット"]
            template_sheet.copy()

            output_book = app.books.active
            sheet = output_book.sheets[0]

            # シート名変更
            sheet.name = "見積書"

            # 他シート削除
            for ws in list(output_book.sheets):
                if ws.name != "見積書":
                    ws.delete()

            # =====================================
            # 基本情報
            # =====================================

            sheet.range("L1").value = data.estimate_no

            # 出力日
            sheet.range("K3").value = datetime.today().strftime("%Y/%m/%d")

            sheet.range("A7").value = self._add_onchu(data.department)

            sheet.range("E14").value = data.subject

            sheet.range("J18").value = self._remove_yen(data.amount)

            # 伝票番号
            sheet.range("B25").value = "伝票ＮＯ．"
            sheet.range("C25").value = data.voucher_no

            # 申請書No
            sheet.range("C26").value = data.application_no

            sheet.range("C30").value = data.due_date

            # J30は文章の日付だけ置換
            original = sheet.range("J30").value

            sheet.range("J30").value = self._replace_date_in_text(
                original,
                data.due_date
            )

            # =====================================
            # 成果物
            # =====================================

            start_row = 18
            end_row = 24

            outputs = [
                x for x in data.outputs
                if x.strip()
            ]

            # A18～A24だけクリア
            for row in range(start_row, end_row + 1):
                sheet.range(f"A{row}").value = None

            for i, output in enumerate(outputs):

                row = start_row + i

                if row > end_row:
                    break

                sheet.range(f"A{row}").value = i + 1
                sheet.range(f"B{row}").value = self._remove_number(output)

            # =====================================
            # 保存
            # =====================================

            output_book.save(output_path)

            # PDF出力
            pdf_path = os.path.splitext(output_path)[0] + ".pdf"

            sheet.api.ExportAsFixedFormat(
                Type=0,
                Filename=pdf_path
            )

        finally:

            try:
                if output_book:
                    output_book.close()
            except Exception:
                pass

            try:
                if template_book:
                    template_book.close()
            except Exception:
                pass

            try:
                app.quit()
            except Exception:
                pass

    # =====================================
    # 御中追加
    # =====================================
    def _add_onchu(self, department):

        department = department.strip()

        if not department:
            return ""

        if department.endswith("御中"):
            return department

        return department + "　御中"

    # =====================================
    # 円削除
    # =====================================
    def _remove_yen(self, amount):

        return amount.replace("円", "").strip()

    # =====================================
    # 成果物番号削除
    # =====================================
    def _remove_number(self, text):

        if not text:
            return ""

        return re.sub(
            r"^[①-⑳]\s*",
            "",
            text
        ).strip()

    # =====================================
    # J30の日付だけ置換
    # =====================================
    def _replace_date_in_text(self, text, due_date):

        if text is None:
            return due_date

        text = str(text)

        pattern = r"\d{4}年\d{1,2}月\d{1,2}日"

        if re.search(pattern, text):
            return re.sub(
                pattern,
                due_date,
                text
            )

        pattern = r"\d{4}/\d{1,2}/\d{1,2}"

        if re.search(pattern, text):
            return re.sub(
                pattern,
                due_date,
                text
            )

        return text