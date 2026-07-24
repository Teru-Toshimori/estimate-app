import os
import shutil
from datetime import datetime

from services.excel.excel_automation_helper import ExcelAutomationSession


class TgExcelWriter:
    """
    TG見積書を出力する。

    入力Excelを出力先へコピーした後、Excel COMを使用して
    セルを書き換える。

    openpyxlは使用しないため、テンプレート内の以下の要素を
    できる限りそのまま保持できる。

    ・決裁欄横の枠や図形
    ・画像
    ・テキストボックス
    ・印刷設定
    ・改ページ
    ・マクロ（xlsm）
    """

    def write(
        self,
        input_excel_path: str,
        output_excel_path: str,
        data: dict,
        estimate_no,
    ) -> None:
        """
        TG見積書Excelを作成する。

        1. 入力Excelを出力先へコピー
        2. Excel COMでコピー先を開く
        3. 必要なセルだけを書き換える
        4. 上書き保存する
        """

        input_excel_path = os.path.abspath(
            input_excel_path
        )
        output_excel_path = os.path.abspath(
            output_excel_path
        )

        if not os.path.exists(
            input_excel_path
        ):
            raise FileNotFoundError(
                "TG見積書の入力Excelが"
                "見つかりません。\n\n"
                f"{input_excel_path}"
            )

        if not os.path.isfile(
            input_excel_path
        ):
            raise ValueError(
                "TG見積書の入力元に"
                "Excelファイルを指定してください。\n\n"
                f"{input_excel_path}"
            )

        output_directory = os.path.dirname(
            output_excel_path
        )

        if output_directory:
            os.makedirs(
                output_directory,
                exist_ok=True,
            )

        # 既存の出力ファイルが開かれている場合は、
        # 削除時に分かりやすいエラーを出す。
        if os.path.exists(
            output_excel_path
        ):
            try:
                os.remove(
                    output_excel_path
                )

            except PermissionError as error:
                raise PermissionError(
                    "既存のTG見積書を"
                    "削除できません。\n"
                    "Excelで開かれていないか"
                    "確認してください。\n\n"
                    f"{output_excel_path}"
                ) from error

        # テンプレート全体をそのままコピーする。
        # 図形・画像・マクロ・印刷設定などを保持するため、
        # openpyxlでは開かない。
        shutil.copy2(
            input_excel_path,
            output_excel_path,
        )

        session = ExcelAutomationSession()
        app = None
        book = None
        sheet = None

        try:
            app = session.start()

            book = app.books.open(
                output_excel_path,
                update_links=False,
                read_only=False,
                ignore_read_only_recommended=True,
                notify=False,
                add_to_mru=False,
            )

            if len(book.sheets) == 0:
                raise ValueError(
                    "TG見積書にシートが"
                    "存在しません。\n\n"
                    f"{output_excel_path}"
                )

            sheet = book.sheets[0]

            # =====================================
            # 基本情報
            # =====================================
            sheet.range("G1").value = (
                self._safe_text(
                    estimate_no
                )
            )

            # 発行日
            sheet.range("G2").value = (
                datetime.today()
            )
            sheet.range("G2").number_format = (
                "yyyy年m月d日"
            )

            # 元セルF6の値をB6へコピー
            sheet.range("B6").value = (
                sheet.range("F6").value
            )

            sheet.range("B8").value = 60
            sheet.range("F8").value = 1

            sheet.range("F4").value = (
                "エイム株式会社"
            )

            sheet.range("F5").value = (
                "齊藤　政輝"
            )

            sheet.range("E17").value = 1

            # =====================================
            # 品名・金額
            # =====================================
            subject = self._safe_text(
                data.get(
                    "品名",
                    "",
                )
            ).strip()

            amount = self.to_int(
                data.get(
                    "金額",
                    "",
                )
            )

            if subject:
                sheet.range("B4").value = (
                    subject
                )
                sheet.range("B17").value = (
                    subject
                )

            if amount:
                sheet.range("B10").value = (
                    amount
                )
                sheet.range("F17").value = (
                    amount
                )

            # =====================================
            # 保存
            # =====================================
            book.save(
                output_excel_path
            )

            if not os.path.exists(
                output_excel_path
            ):
                raise RuntimeError(
                    "TG見積書Excelが"
                    "作成されませんでした。\n\n"
                    f"{output_excel_path}"
                )

        except Exception:
            # 処理失敗時に不完全な出力ファイルが
            # 残らないよう、ブックを閉じた後に削除する。
            raise

        finally:
            if book is not None:
                try:
                    book.close()

                except Exception:
                    pass

            sheet = None
            book = None
            app = None

            session.close()

    def export_pdf(
        self,
        excel_path: str,
        pdf_path: str,
    ) -> None:
        """
        作成済みTG見積書ExcelをPDFへ出力する。

        Excel COMで直接PDF化するため、
        決裁欄横の図形や画像もPDFへ反映される。
        """

        excel_path = os.path.abspath(
            excel_path
        )
        pdf_path = os.path.abspath(
            pdf_path
        )

        if not os.path.exists(
            excel_path
        ):
            raise FileNotFoundError(
                "PDF出力元のTG見積書Excelが"
                "見つかりません。\n\n"
                f"{excel_path}"
            )

        pdf_directory = os.path.dirname(
            pdf_path
        )

        if pdf_directory:
            os.makedirs(
                pdf_directory,
                exist_ok=True,
            )

        if os.path.exists(
            pdf_path
        ):
            try:
                os.remove(
                    pdf_path
                )

            except PermissionError as error:
                raise PermissionError(
                    "既存のTG見積書PDFを"
                    "削除できません。\n"
                    "PDFビューアーで"
                    "開かれていないか"
                    "確認してください。\n\n"
                    f"{pdf_path}"
                ) from error

        session = ExcelAutomationSession()
        app = None
        book = None

        try:
            app = session.start()

            book = app.books.open(
                excel_path,
                update_links=False,
                read_only=True,
                ignore_read_only_recommended=True,
                notify=False,
                add_to_mru=False,
            )

            session.enable_print_communication()

            # ブック全体をPDF化する。
            # テンプレート側の印刷範囲・改ページ・図形を保持する。
            book.api.ExportAsFixedFormat(
                Type=0,
                Filename=pdf_path,
                Quality=0,
                IncludeDocProperties=True,
                IgnorePrintAreas=False,
                OpenAfterPublish=False,
            )

            if not os.path.exists(
                pdf_path
            ):
                raise RuntimeError(
                    "TG見積書PDFが"
                    "作成されませんでした。\n\n"
                    f"{pdf_path}"
                )

        finally:
            if book is not None:
                try:
                    book.close()

                except Exception:
                    pass

            book = None
            app = None

            session.close()

    @staticmethod
    def to_int(
        value,
    ) -> int:
        """
        金額文字列を整数へ変換する。

        対応例:
            1,234,567円
            ￥1,234,567
            1234567.0
        """

        if value in (
            None,
            "",
        ):
            return 0

        text = (
            str(value)
            .replace(
                ",",
                "",
            )
            .replace(
                "円",
                "",
            )
            .replace(
                "￥",
                "",
            )
            .replace(
                "¥",
                "",
            )
            .strip()
        )

        if not text:
            return 0

        try:
            return int(
                float(text)
            )

        except ValueError as error:
            raise ValueError(
                "TG見積書の金額を"
                "数値へ変換できません。\n\n"
                f"入力値：{value}"
            ) from error

    @staticmethod
    def _safe_text(
        value,
    ) -> str:
        if value is None:
            return ""

        return str(
            value
        )
