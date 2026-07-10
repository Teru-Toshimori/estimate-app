from copy import copy
from datetime import datetime
import getpass
import unicodedata

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font


class LedgerWriter:
    """
    管理台帳へ新しい行を追加するクラス。

    主な処理:
    - 全シートのC列から依頼部署を検索
    - B列とD列を前行から連番で採番
    - PDF解析結果を管理台帳へ書き込み
    - 前行の書式をB列～L列へコピー
    - 必要なセル書式を上書き
    """

    def write(self, excel_path: str, data) -> dict:
        """
        管理台帳へ1行追加する。

        Args:
            excel_path: 管理台帳Excelのパス
            data: EstimateData

        Returns:
            台帳記入結果を格納した辞書
        """

        workbook = load_workbook(
            excel_path,
            keep_links=False,
        )

        sheet = None

        try:
            sheet = self.find_sheet_by_department(
                workbook,
                data.department,
            )

            if sheet is None:
                raise ValueError(
                    "依頼部署に該当するシートが見つかりません。\n"
                    f"依頼部署：{data.department}"
                )

            last_row = self.find_last_row(sheet)
            new_row = last_row + 1

            previous_no = sheet[f"B{last_row}"].value
            previous_estimate_no = sheet[f"D{last_row}"].value

            new_no = self.to_int(previous_no) + 1
            new_estimate_no = self.to_int(previous_estimate_no) + 1

            today = datetime.today()
            estimate_month = self.get_next_month_text(today)
            user_name = getpass.getuser()

            # ITK14642 のような値から数値部分だけ取り出す
            application_number = self.to_int(data.application_no)

            # 前行の書式・行高を新しい行へコピー
            self.copy_row_style(
                sheet=sheet,
                source_row=last_row,
                target_row=new_row,
                start_column=2,   # B列
                end_column=12,   # L列
            )

            # =============================
            # 台帳へ値を書き込み
            # =============================
            sheet[f"B{new_row}"] = new_no
            sheet[f"C{new_row}"] = data.department
            sheet[f"D{new_row}"] = new_estimate_no
            sheet[f"F{new_row}"] = application_number
            sheet[f"G{new_row}"] = data.model_code
            sheet[f"I{new_row}"] = estimate_month
            sheet[f"J{new_row}"] = today.date()
            sheet[f"K{new_row}"] = user_name

            # =============================
            # フォント設定
            # B列～L列
            # =============================
            for column in range(2, 13):
                cell = sheet.cell(
                    row=new_row,
                    column=column,
                )

                cell.font = Font(
                    name="Meiryo UI",
                    size=10,
                    bold=cell.font.bold,
                    italic=cell.font.italic,
                    vertAlign=cell.font.vertAlign,
                    underline=cell.font.underline,
                    strike=cell.font.strike,
                    color=copy(cell.font.color),
                )

            # =============================
            # 配置設定
            # =============================
            # C列・G列を上下中央にする
            for column_letter in ("C", "G"):

                current = sheet[f"{column_letter}{new_row}"].alignment

                sheet[f"{column_letter}{new_row}"].alignment = Alignment(
                    horizontal=current.horizontal,   # 元の左右配置を維持
                    vertical="center",               # 上下中央のみ
                    wrap_text=current.wrap_text,
                    shrink_to_fit=current.shrink_to_fit,
                    text_rotation=current.text_rotation,
                )

            # その他の入力セルも上下中央、左右中央にする
            for column_letter in ("B", "D", "F", "I", "J", "K"):
                current = sheet[f"{column_letter}{new_row}"].alignment

                sheet[f"{column_letter}{new_row}"].alignment = Alignment(
                    horizontal="center",
                    vertical="center",
                    wrap_text=current.wrap_text,
                    shrink_to_fit=current.shrink_to_fit,
                    text_rotation=current.text_rotation,
                )

            # =============================
            # 数値・日付書式
            # =============================
            # F列：数値のまま保持し、表示時にITKを付ける
            sheet[f"F{new_row}"].number_format = '"ITK"0'

            # J列：日付のみ表示し、時間は表示しない
            sheet[f"J{new_row}"].number_format = "yyyy/mm/dd"

            # =============================
            # 列幅
            # =============================
            # B列の値に合わせて必要な幅へ調整
            self.adjust_column_width(
                sheet=sheet,
                column_letter="B",
                minimum_width=4,
                maximum_width=12,
            )

            workbook.save(excel_path)

            return {
                "sheet_name": sheet.title,
                "row": new_row,
                "estimate_no": str(new_estimate_no),
                "issue_date": today.strftime("%Y/%m/%d"),
                "user_name": user_name,
            }

        finally:
            workbook.close()

    def find_sheet_by_department(self, workbook, department: str):
        """
        全シートのC列を検索し、依頼部署が含まれるシートを返す。

        全角・半角、空白の違いは正規化して比較する。
        """

        normalized_department = self.normalize_text(department)

        if not normalized_department:
            return None

        for sheet in workbook.worksheets:
            for cell in sheet["C"]:
                if cell.value is None:
                    continue

                normalized_value = self.normalize_text(cell.value)

                if (
                    normalized_value == normalized_department
                    or normalized_department in normalized_value
                    or normalized_value in normalized_department
                ):
                    return sheet

        return None

    def normalize_text(self, text) -> str:
        """
        文字列を検索比較用に正規化する。

        例:
        - 第３シート設計部 → 第3シート設計部
        - 全角スペース、半角スペースを削除
        """

        if text is None:
            return ""

        normalized = unicodedata.normalize(
            "NFKC",
            str(text).strip(),
        )

        normalized = normalized.replace("　", "")
        normalized = normalized.replace(" ", "")

        return normalized

    def find_last_row(self, sheet) -> int:
        """
        B列を基準に最終データ行を取得する。
        """

        for row_number in range(
            sheet.max_row,
            1,
            -1,
        ):
            if sheet[f"B{row_number}"].value is not None:
                return row_number

        return 1

    def to_int(self, value) -> int:
        """
        値から数値部分を取得してintへ変換する。

        例:
        - 12 → 12
        - "12" → 12
        - "ITK14642" → 14642
        """

        if value is None:
            return 0

        text = str(value).strip()

        try:
            return int(text)
        except (TypeError, ValueError):
            digits = "".join(
                character
                for character in text
                if character.isdigit()
            )

            return int(digits) if digits else 0

    def get_next_month_text(self, date_value: datetime) -> str:
        """
        操作日の翌月を「8月」の形式で返す。
        """

        next_month = date_value.month + 1

        if next_month == 13:
            next_month = 1

        return f"{next_month}月"

    def copy_row_style(
        self,
        sheet,
        source_row: int,
        target_row: int,
        start_column: int,
        end_column: int,
    ) -> None:
        """
        指定範囲の前行書式を新しい行へコピーする。

        コピー対象:
        - フォント
        - 塗りつぶし
        - 罫線
        - 配置
        - 表示形式
        - 保護
        - 行高
        """

        source_height = sheet.row_dimensions[source_row].height

        if source_height is not None:
            sheet.row_dimensions[target_row].height = source_height

        for column in range(
            start_column,
            end_column + 1,
        ):
            source_cell = sheet.cell(
                row=source_row,
                column=column,
            )
            target_cell = sheet.cell(
                row=target_row,
                column=column,
            )

            if source_cell.has_style:
                target_cell.font = copy(source_cell.font)
                target_cell.fill = copy(source_cell.fill)
                target_cell.border = copy(source_cell.border)
                target_cell.alignment = copy(source_cell.alignment)
                target_cell.number_format = source_cell.number_format
                target_cell.protection = copy(source_cell.protection)

    def adjust_column_width(
        self,
        sheet,
        column_letter: str,
        minimum_width: float = 4,
        maximum_width: float = 50,
    ) -> None:
        """
        指定列の文字数に合わせて列幅を調整する。
        """

        maximum_length = 0

        for cell in sheet[column_letter]:
            if cell.value is None:
                continue

            text_length = len(str(cell.value))

            if text_length > maximum_length:
                maximum_length = text_length

        calculated_width = maximum_length + 2
        calculated_width = max(
            minimum_width,
            calculated_width,
        )
        calculated_width = min(
            maximum_width,
            calculated_width,
        )

        sheet.column_dimensions[column_letter].width = calculated_width