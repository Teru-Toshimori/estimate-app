import os
import shutil
import logging
from datetime import datetime
import win32com.client

import openpyxl

from models.tg_estimate_data import TgEstimateData


logger = logging.getLogger(__name__)


class TgExcelWriter:
    """
    TG見積書Excel出力

    処理
        テンプレートコピー
            ↓
        必要セルへ値設定
            ↓
        保存
    """

    TEMPLATE_PATH = "resources/TG_見積書フォーマット.xlsm"

    def write(
        self,
        template_path: str,
        output_path: str,
        data: TgEstimateData,
    ) -> None:
        """
        Excel出力

        Args:
            template_path:
                Excelテンプレート

            output_path:
                保存先

            data:
                OCR結果
        """

        logger.info("Excel出力開始")

        # -----------------------------
        # テンプレート存在確認
        # -----------------------------
        if not os.path.exists(template_path):
            raise FileNotFoundError(
                f"テンプレートがありません。\n{template_path}"
            )

        # -----------------------------
        # テンプレートコピー
        # -----------------------------
        shutil.copy2(
            template_path,
            output_path,
        )

        logger.info(
            "テンプレートコピー完了 : %s",
            output_path
        )

        # -----------------------------
        # VBA保持で開く
        # -----------------------------
        workbook = openpyxl.load_workbook(
            output_path,
            keep_vba=True,
        )

        sheet = workbook.active

        # ===================================
        # 固定値
        # ===================================

        sheet["G1"] = "14101201"

        sheet["B6"] = "QY6247"
        sheet["F6"] = "QY6247"

        sheet["B8"] = 60
        sheet["F8"] = 1

        sheet["F4"] = "エイム株式会社"
        sheet["F5"] = "齊藤　政輝"

        sheet["E17"] = 1

        # ===================================
        # 今日の日付
        # ===================================

        sheet["G2"] = datetime.now().strftime(
            "%Y年%m月%d日"
        )

        # ===================================
        # OCR結果
        # ===================================

        if data.subject:
            sheet["B4"] = data.subject
            sheet["B17"] = data.subject

        if data.amount:
            amount = int(
                str(data.amount).replace(",", "")
            )

            sheet["B10"] = amount
            sheet["F17"] = amount

        # -----------------------------
        # 保存
        # -----------------------------
        workbook.save(output_path)
        workbook.close()

        logger.info(
            "Excel保存完了 : %s",
            output_path
        )

    def export_pdf(
        self,
        excel_path: str,
        pdf_path: str,
    ) -> None:
        """
        ExcelをPDFへ変換
        """

        logger.info("PDF出力開始")

        excel = win32com.client.DispatchEx(
            "Excel.Application"
        )

        excel.Visible = False
        excel.DisplayAlerts = False

        try:

            workbook = excel.Workbooks.Open(
                os.path.abspath(excel_path)
            )

            workbook.ExportAsFixedFormat(
                0,      # PDF
                os.path.abspath(pdf_path)
            )

            workbook.Close(False)

            logger.info(
                "PDF保存完了 : %s",
                pdf_path
            )

        finally:

            excel.Quit()