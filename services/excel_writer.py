import gc
import os
import re
from datetime import datetime
from pathlib import Path

import xlwings as xw
import win32print


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

        original_default_printer = None
        selected_pdf_printer = None

        app = None
        template_book = None
        output_book = None
        sheet = None
        print_communication_changed = False

        try:
            # Excel起動前に、利用可能なPDFプリンターを既定へ一時設定する。
            # 既定プリンターがオフラインの場合に表示される
            # 「プリンターの接続を待っています」を防ぐ。
            original_default_printer, selected_pdf_printer = (
                self._prepare_pdf_printer()
            )

            app = xw.App(visible=False, add_book=False)
            self._configure_excel_application(app)
            self._set_excel_active_printer(app, selected_pdf_printer)

            # PageSetup変更時などに発生するプリンター通信を一時停止する。
            # Excelや環境によって未対応の場合があるため、安全に設定する。
            print_communication_changed = self._set_print_communication(
                app,
                enabled=False,
            )

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

            # PDF出力前に、保留していたプリンター通信を有効に戻す。
            # ExportAsFixedFormatは印刷ダイアログを使わず直接PDF化する。
            if print_communication_changed:
                self._set_print_communication(app, enabled=True)
                print_communication_changed = False

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
            if app is not None and print_communication_changed:
                self._set_print_communication(app, enabled=True)

            self._close_book(output_book)
            self._close_book(template_book)

            sheet = None
            output_book = None
            template_book = None

            if app is not None:
                try:
                    app.display_alerts = False
                    app.screen_updating = False
                except Exception:
                    pass

                try:
                    app.quit()
                except Exception:
                    try:
                        app.kill()
                    except Exception:
                        pass

            app = None

            # アプリ実行前の既定プリンターへ戻す。
            self._restore_default_printer(original_default_printer)

            gc.collect()


    # =====================================
    # プリンター設定
    # =====================================
    def _prepare_pdf_printer(self) -> tuple[str | None, str]:
        """
        Excel起動前に利用可能なPDFプリンターを既定へ一時設定する。

        戻り値:
            (元の既定プリンター名, 使用するPDFプリンター名)
        """
        try:
            original_printer = win32print.GetDefaultPrinter()
        except Exception:
            original_printer = None

        available_printers = {
            printer_info[2]
            for printer_info in win32print.EnumPrinters(
                win32print.PRINTER_ENUM_LOCAL
                | win32print.PRINTER_ENUM_CONNECTIONS
            )
        }

        preferred_printers = (
            "Microsoft Print to PDF",
            "Microsoft XPS Document Writer",
        )

        selected_printer = next(
            (
                printer_name
                for printer_name in preferred_printers
                if printer_name in available_printers
            ),
            None,
        )

        if selected_printer is None:
            raise RuntimeError(
                "PDF出力に使用できるプリンターが見つかりません。\n\n"
                "Windowsの『Microsoft Print to PDF』を有効にしてください。"
            )

        try:
            win32print.SetDefaultPrinter(selected_printer)
        except Exception as error:
            raise RuntimeError(
                "PDF出力用プリンターを一時設定できませんでした。\n\n"
                f"プリンター：{selected_printer}"
            ) from error

        return original_printer, selected_printer

    def _set_excel_active_printer(self, app, printer_name: str | None) -> None:
        """Excel側でも使用プリンターを明示する。"""
        if not printer_name:
            return

        # ActivePrinterは環境によってポート名を含む表記を要求する。
        # 既定プリンターは先に切り替えてあるため、失敗しても処理は継続する。
        try:
            app.api.ActivePrinter = printer_name
            return
        except Exception:
            pass

        try:
            printer_handle = win32print.OpenPrinter(printer_name)
            try:
                printer_info = win32print.GetPrinter(printer_handle, 2)
                port_name = printer_info.get("pPortName", "")
            finally:
                win32print.ClosePrinter(printer_handle)

            if port_name:
                app.api.ActivePrinter = f"{printer_name} on {port_name}"
        except Exception:
            pass

    def _restore_default_printer(self, printer_name: str | None) -> None:
        """処理前の既定プリンターへ戻す。"""
        if not printer_name:
            return

        try:
            win32print.SetDefaultPrinter(printer_name)
        except Exception:
            # 復元失敗は見積書作成自体のエラーにはしない。
            pass


    # =====================================
    # Excelアプリケーション設定
    # =====================================
    def _configure_excel_application(self, app) -> None:
        """
        自動処理中の画面表示、確認メッセージ、リンク更新、
        マクロ、イベントなどを可能な範囲で抑止する。
        """
        app.visible = False
        app.display_alerts = False
        app.screen_updating = False

        api_settings = {
            "DisplayAlerts": False,
            "ScreenUpdating": False,
            "EnableEvents": False,
            "AskToUpdateLinks": False,
            "Interactive": False,
            # msoAutomationSecurityForceDisable
            "AutomationSecurity": 3,
        }

        for property_name, value in api_settings.items():
            try:
                setattr(app.api, property_name, value)
            except Exception:
                # Excelのバージョンや環境によって利用できない
                # プロパティがあるため、その設定だけをスキップする。
                pass

    def _set_print_communication(self, app, enabled: bool) -> bool:
        """
        Excelのプリンター通信を切り替える。

        設定に成功した場合はTrue、未対応・失敗時はFalseを返す。
        """
        try:
            app.api.PrintCommunication = enabled
            return True
        except Exception:
            return False

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
