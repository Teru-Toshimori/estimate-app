import os
import re
import xlwings as xw


class ExcelWriter:

    def write(self, template_path: str, output_path: str, data):

        template_path = os.path.abspath(template_path)
        output_path = os.path.abspath(output_path)

        app = xw.App(visible=False, add_book=False)

        template_book = None
        output_book = None

        try:
            template_book = app.books.open(template_path)

            # 「フォーマット」シートだけ新規ブックへコピー
            template_sheet = template_book.sheets["フォーマット"]
            template_sheet.copy()

            output_book = app.books.active
            sheet = output_book.sheets[0]

            # シート名を変更
            sheet.name = "見積書"

            # 念のため、出力ブック内の他シートを削除
            for ws in list(output_book.sheets):
                if ws.name != "見積書":
                    ws.delete()

            # 値の加工
            department = self._add_onchu(data.department)
            amount = self._remove_yen(data.amount)
            due_date = data.due_date

            # 転記
            sheet.range("L1").value = data.estimate_no
            sheet.range("K3").value = data.issue_date

            sheet.range("A7").value = department
            sheet.range("E14").value = data.subject

            sheet.range("J18").value = amount
            sheet.range("C26").value = data.application_no

            sheet.range("C30").value = due_date

            # J30は既存文の「日付部分だけ」置換
            original_j30 = sheet.range("J30").value
            sheet.range("J30").value = self._replace_date_in_text(
                original_j30,
                due_date
            )

            # 成果物 B18以降
            for i, output in enumerate(data.outputs):
                sheet.range(f"B{18 + i}").value = self._remove_number(output)

            output_book.save(output_path)

        finally:
            try:
                if output_book is not None:
                    output_book.close()
            except Exception:
                pass

            try:
                if template_book is not None:
                    template_book.close()
            except Exception:
                pass

            try:
                app.quit()
            except Exception:
                pass

    def _add_onchu(self, department: str) -> str:
        department = department.strip()

        if not department:
            return ""

        if department.endswith("御中"):
            return department

        return department + "　御中"

    def _remove_yen(self, amount: str) -> str:
        return amount.replace("円", "").strip()

    def _replace_date_in_text(self, text, due_date: str) -> str:
        if text is None:
            return due_date

        text = str(text)

        # 例：2025年01月24日 を置換
        pattern = r"\d{4}年\d{1,2}月\d{1,2}日"

        if re.search(pattern, text):
            return re.sub(pattern, due_date, text)

        # 例：2025/01/24 を置換
        pattern_slash = r"\d{4}/\d{1,2}/\d{1,2}"

        if re.search(pattern_slash, text):
            return re.sub(pattern_slash, due_date, text)

        # 日付が見つからない場合は末尾に追記
        return text
    
    def _remove_number(self, text: str) -> str:
        """
        成果物名称の先頭の番号（①～⑳など）を削除する
        """

        if not text:
            return ""

        import re

        # 先頭の丸数字と後ろの空白を削除
        return re.sub(r"^[①-⑳]\s*", "", text).strip()