import os
import re
import xlwings as xw


class ExcelWriter:

    def write(self, template_path: str, output_path: str, data):

        template_path = os.path.abspath(template_path)
        output_path = os.path.abspath(output_path)

        # 同名ファイルが存在する場合は削除
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

            # 「フォーマット」シートのみ新規ブックへコピー
            template_sheet = template_book.sheets["フォーマット"]
            template_sheet.copy()

            output_book = app.books.active
            sheet = output_book.sheets[0]

            # シート名変更
            sheet.name = "見積書"

            # 不要シート削除
            for ws in list(output_book.sheets):
                if ws.name != "見積書":
                    ws.delete()

            # =============================
            # 基本情報
            # =============================
            sheet.range("L1").value = data.estimate_no
            sheet.range("K3").value = data.issue_date

            sheet.range("A7").value = self._add_onchu(data.department)
            sheet.range("E14").value = data.subject

            sheet.range("J18").value = self._remove_yen(data.amount)

            sheet.range("C26").value = data.application_no

            sheet.range("C30").value = data.due_date

            # J30は文章の日付部分だけ置換
            original_j30 = sheet.range("J30").value
            sheet.range("J30").value = self._replace_date_in_text(
                original_j30,
                data.due_date
            )

            # =============================
            # 成果物
            # =============================
            start_row = 18
            end_row = 24

            outputs = [
                output
                for output in data.outputs
                if output.strip()
            ]

            # A18～A24 のNoだけクリア
            # （B列は触らない）
            for row in range(start_row, end_row + 1):
                sheet.range(f"A{row}").value = None

            # 成果物を書き込み
            for i, output in enumerate(outputs):

                row = start_row + i

                if row > end_row:
                    break

                # No
                sheet.range(f"A{row}").value = i + 1

                # 成果物名称
                sheet.range(f"B{row}").value = self._remove_number(output)

            # 保存
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

    # ======================================
    # 依頼部署へ「御中」を付加
    # ======================================
    def _add_onchu(self, department: str) -> str:

        department = department.strip()

        if not department:
            return ""

        if department.endswith("御中"):
            return department

        return department + "　御中"

    # ======================================
    # 円を削除
    # ======================================
    def _remove_yen(self, amount: str) -> str:

        return amount.replace("円", "").strip()

    # ======================================
    # 成果物番号削除
    # ① ② ③ …
    # ======================================
    def _remove_number(self, text: str) -> str:

        if not text:
            return ""

        return re.sub(
            r"^[①-⑳]\s*",
            "",
            text
        ).strip()

    # ======================================
    # J30の日付だけ置換
    # ======================================
    def _replace_date_in_text(self, text, due_date: str) -> str:

        if text is None:
            return due_date

        text = str(text)

        # 2025年01月24日
        pattern = r"\d{4}年\d{1,2}月\d{1,2}日"

        if re.search(pattern, text):
            return re.sub(
                pattern,
                due_date,
                text
            )

        # 2025/01/24
        pattern = r"\d{4}/\d{1,2}/\d{1,2}"

        if re.search(pattern, text):
            return re.sub(
                pattern,
                due_date,
                text
            )

        return text