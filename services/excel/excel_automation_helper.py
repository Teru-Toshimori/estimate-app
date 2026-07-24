import gc

import win32print
import xlwings as xw


class ExcelAutomationSession:
    """
    Excelの自動操作を安全に実行するための共通セッション。

    特調TBと同様に、Excel起動前に利用可能なPDFプリンターを
    既定プリンターへ一時設定し、警告・イベント・リンク更新・
    プリンター通信を可能な範囲で抑止する。
    """

    def __init__(self):
        self.app = None
        self.original_default_printer = None
        self.selected_pdf_printer = None
        self.print_communication_changed = False

    def start(self):
        self.original_default_printer, self.selected_pdf_printer = (
            self._prepare_pdf_printer()
        )

        self.app = xw.App(visible=False, add_book=False)
        self._configure_excel_application()
        self._set_excel_active_printer()

        self.print_communication_changed = self._set_print_communication(
            enabled=False
        )

        return self.app

    def enable_print_communication(self) -> None:
        if self.app is None or not self.print_communication_changed:
            return

        self._set_print_communication(enabled=True)
        self.print_communication_changed = False

    def close(self) -> None:
        if self.app is not None and self.print_communication_changed:
            self._set_print_communication(enabled=True)
            self.print_communication_changed = False

        if self.app is not None:
            try:
                self.app.display_alerts = False
                self.app.screen_updating = False
            except Exception:
                pass

            try:
                self.app.quit()
            except Exception:
                try:
                    self.app.kill()
                except Exception:
                    pass

        self.app = None
        self._restore_default_printer(self.original_default_printer)
        gc.collect()

    def _prepare_pdf_printer(self) -> tuple[str | None, str]:
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
                "Excel処理に使用できるプリンターが見つかりません。\n\n"
                "Windowsの『Microsoft Print to PDF』を有効にしてください。"
            )

        try:
            win32print.SetDefaultPrinter(selected_printer)
        except Exception as error:
            raise RuntimeError(
                "Excel処理用プリンターを一時設定できませんでした。\n\n"
                f"プリンター：{selected_printer}"
            ) from error

        return original_printer, selected_printer

    def _set_excel_active_printer(self) -> None:
        if self.app is None or not self.selected_pdf_printer:
            return

        try:
            self.app.api.ActivePrinter = self.selected_pdf_printer
            return
        except Exception:
            pass

        try:
            printer_handle = win32print.OpenPrinter(
                self.selected_pdf_printer
            )
            try:
                printer_info = win32print.GetPrinter(printer_handle, 2)
                port_name = printer_info.get("pPortName", "")
            finally:
                win32print.ClosePrinter(printer_handle)

            if port_name:
                self.app.api.ActivePrinter = (
                    f"{self.selected_pdf_printer} on {port_name}"
                )
        except Exception:
            pass

    def _restore_default_printer(self, printer_name: str | None) -> None:
        if not printer_name:
            return

        try:
            win32print.SetDefaultPrinter(printer_name)
        except Exception:
            pass

    def _configure_excel_application(self) -> None:
        if self.app is None:
            return

        self.app.visible = False
        self.app.display_alerts = False
        self.app.screen_updating = False

        api_settings = {
            "DisplayAlerts": False,
            "ScreenUpdating": False,
            "EnableEvents": False,
            "AskToUpdateLinks": False,
            "Interactive": False,
            "AutomationSecurity": 3,
        }

        for property_name, value in api_settings.items():
            try:
                setattr(self.app.api, property_name, value)
            except Exception:
                pass

    def _set_print_communication(self, enabled: bool) -> bool:
        if self.app is None:
            return False

        try:
            self.app.api.PrintCommunication = enabled
            return True
        except Exception:
            return False
