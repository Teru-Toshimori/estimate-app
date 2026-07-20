import os
import re
from datetime import datetime
from pathlib import Path

from services.excel_automation_helper import ExcelAutomationSession


class ExcelWriter:
    """
    見積書テンプレートへデータを書き込み、Excel・PDFを出力する。

    Excelはバックグラウンドで起動し、警告表示、リンク更新、
    イベント実行、画面更新、不要なプリンター通信を抑制する。
    """

    def write(self, template_path: str, output_path: str, data) -> None:
        template_path = str(Path(template_path).resolve())
        output_path = str(Path(output_path).resolve())
        pdf_path = str(Path(output_path).with_suffix(".pdf"))

        self._validate_template(template_path)
        Path(output_path).parent.mkdir(parents=True, exist_ok=True)

        self._remove_existing_file(output_path)
        self._remove_existing_file(pdf_path)

        session = ExcelAutomationSession()
        app = None
        template_book = None
        output_book = None
        sheet = None

        try:
            # Excel起動、PDFプリンターの一時設定、警告・イベント・
            # リンク更新・プリンター通信の抑止を共通セッションへ任せる。
            app = session.start()

            # テンプレートを読み取り専用で開く。
            # リンク更新、読み取り専用推奨、通知、最近使ったファイルへの追加を抑止する。
            template_book = app.books.open(
                template_path,
                update_links=False,
                read_only=True,
                ignore_read_only_recommended=True,
                notify=False,
                add_to_mru=False,
            )

            if "フォーマット" not in [ws.name for ws in template_book.sheets]:
                raise ValueError(
                    "テンプレートに「フォーマット」シートが見つかりません。\n\n"
                    f"{template_path}"
                )

            # 「フォーマット」シートを新規ブックへコピーする。
            template_sheet = template_book.sheets["フォーマット"]
            template_sheet.copy()

            output_book = app.books.active
            sheet = output_book.sheets[0]

            # シート名変更
            sheet.name = "見積書"

            # 他シート削除
            for worksheet in list(output_book.sheets):
                if worksheet.name != "見積書":
                    worksheet.delete()

            # =====================================
            # 基本情報
            # =====================================
            sheet.range("L1").value = self._safe_text(data.estimate_no)

            # 出力日
            sheet.range("K3").value = datetime.today().strftime("%Y/%m/%d")

            sheet.range("A7").value = self._add_onchu(data.department)
            sheet.range("E14").value = self._safe_text(data.subject)
            sheet.range("J18").value = self._remove_yen(data.amount)

            # 伝票番号
            sheet.range("B25").value = "伝票ＮＯ．"
            sheet.range("C25").value = self._safe_text(data.voucher_no)

            # 申請書No
            sheet.range("C26").value = self._safe_text(data.application_no)

            sheet.range("C30").value = self._safe_text(data.due_date)

            # J30は文章の日付だけ置換
            original = sheet.range("J30").value
            sheet.range("J30").value = self._replace_date_in_text(
                original,
                data.due_date,
            )

            # =====================================
            # 成果物
            # =====================================
            start_row = 18
            end_row = 24

            raw_outputs = getattr(data, "outputs", []) or []
            outputs = [
                str(value).strip()
                for value in raw_outputs
                if value is not None and str(value).strip()
            ]

            # A18～A24とB18～B24の旧データをクリアする。
            for row in range(start_row, end_row + 1):
                sheet.range(f"A{row}").value = None
                sheet.range(f"B{row}").value = None

            for index, output in enumerate(outputs):
                row = start_row + index

                if row > end_row:
                    break

                sheet.range(f"A{row}").value = index + 1
                sheet.range(f"B{row}").value = self._remove_number(output)

            # =====================================
            # 保存・PDF出力
            # =====================================
            output_book.save(output_path)

            if not os.path.exists(output_path):
                raise RuntimeError(
                    "見積書Excelが作成されませんでした。\n\n"
                    f"{output_path}"
                )

            # PDF出力前にプリンター通信を有効に戻す。
            # ExportAsFixedFormatは印刷ダイアログを使わず直接PDF化する。
            session.enable_print_communication()

            sheet.api.ExportAsFixedFormat(
                Type=0,
                Filename=pdf_path,
                Quality=0,
                IncludeDocProperties=True,
                IgnorePrintAreas=False,
                OpenAfterPublish=False,
            )

            if not os.path.exists(pdf_path):
                raise RuntimeError(
                    "見積書PDFが作成されませんでした。\n\n"
                    f"{pdf_path}"
                )

        finally:
            # ブックを先に閉じてからExcelセッションを終了する。
            self._close_book(output_book)
            self._close_book(template_book)

            sheet = None
            output_book = None
            template_book = None
            app = None

            # Excel終了と既定プリンターの復元を必ず行う。
            session.close()

    # =====================================
    # ファイル・ブック操作
    # =====================================
    def _validate_template(self, template_path: str) -> None:
        if not os.path.exists(template_path):
            raise FileNotFoundError(
                "見積書テンプレートが見つかりません。\n\n"
                f"{template_path}"
            )

        if not os.path.isfile(template_path):
            raise ValueError(
                "見積書テンプレートにファイルを指定してください。\n\n"
                f"{template_path}"
            )

    def _remove_existing_file(self, file_path: str) -> None:
        if not os.path.exists(file_path):
            return

        try:
            os.remove(file_path)
        except PermissionError as error:
            raise PermissionError(
                "既存の出力ファイルを削除できません。\n"
                "ExcelまたはPDFビューアーで開かれていないか確認してください。\n\n"
                f"{file_path}"
            ) from error

    def _close_book(self, book) -> None:
        if book is None:
            return

        try:
            book.close()
        except Exception:
            pass

    # =====================================
    # 御中追加
    # =====================================
    def _add_onchu(self, department) -> str:
        department_text = self._safe_text(department).strip()

        if not department_text:
            return ""

        if department_text.endswith("御中"):
            return department_text

        return department_text + "　御中"

    # =====================================
    # 円削除
    # =====================================
    def _remove_yen(self, amount) -> str:
        amount_text = self._safe_text(amount)

        return (
            amount_text.replace("円", "")
            .replace("￥", "")
            .replace("¥", "")
            .strip()
        )

    # =====================================
    # 成果物番号削除
    # =====================================
    def _remove_number(self, text) -> str:
        if text is None:
            return ""

        return re.sub(
            r"^[①-⑳]\s*",
            "",
            str(text),
        ).strip()

    # =====================================
    # J30の日付だけ置換
    # =====================================
    def _replace_date_in_text(self, text, due_date) -> str:
        due_date_text = self._safe_text(due_date)

        if text is None:
            return due_date_text

        text = str(text)

        patterns = (
            r"\d{4}年\d{1,2}月\d{1,2}日",
            r"\d{4}/\d{1,2}/\d{1,2}",
            r"\d{4}-\d{1,2}-\d{1,2}",
        )

        for pattern in patterns:
            if re.search(pattern, text):
                return re.sub(
                    pattern,
                    due_date_text,
                    text,
                    count=1,
                )

        return text

    def _safe_text(self, value) -> str:
        if value is None:
            return ""

        return str(value)
