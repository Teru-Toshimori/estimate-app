import re
from pathlib import Path

from services.excel.template_resolver import TemplateResolver
from services.writers.estimate_excel_writer import ExcelWriter
from services.text.work_item_shortener import WorkItemShortener


class TokuchoOtherExcelWriter:
    """
    特調TB以外の見積書を出力する。

    使用テンプレート:
        resources/TB_見積書フォーマット.xlsx

    出力ファイル名:
        部署名_作業項目.xlsx
        部署名_作業項目.pdf

    実際のExcel転記とPDF変換は、
    既存のExcelWriterへ委譲する。
    """

    MAX_FILE_NAME_LENGTH = 150

    def __init__(self):
        self.excel_writer = ExcelWriter()
        self.work_item_shortener = WorkItemShortener()

    # =====================================
    # Excel・PDF出力
    # =====================================
    def write(
        self,
        output_folder: str,
        data,
    ) -> dict:
        """
        特調TB以外の見積書をExcel・PDFで出力する。

        Args:
            output_folder:
                出力先フォルダ

            data:
                PDF抽出結果と採番結果を保持する
                EstimateDataオブジェクト

        Returns:
            {
                "file_name_base": str,
                "excel_path": str,
                "pdf_path": str,
            }
        """

        output_directory = Path(
            output_folder
        )

        output_directory.mkdir(
            parents=True,
            exist_ok=True,
        )

        template_path = TemplateResolver.resolve("tb")

        self.prepare_data(
            data
        )

        file_name_base = (
            self.create_output_file_name(
                department=self.get_department(
                    data
                ),
                work_item=self.get_subject(
                    data
                ),
            )
        )

        # 既存の同名ファイルは削除・上書きせず、
        # 「_2」「_3」…の連番を付けて新規保存する。
        file_name_base = self.create_unique_file_name_base(
            output_directory=output_directory,
            file_name_base=file_name_base,
        )

        output_excel_path = (
            output_directory
            / f"{file_name_base}.xlsx"
        )

        output_pdf_path = (
            output_directory
            / f"{file_name_base}.pdf"
        )

        self.excel_writer.write(
            str(template_path),
            str(output_excel_path),
            data,
        )

        if not output_excel_path.exists():
            raise RuntimeError(
                "見積書Excelが"
                "作成されませんでした。\n\n"
                f"{output_excel_path}"
            )

        if not output_pdf_path.exists():
            raise RuntimeError(
                "見積書PDFが"
                "作成されませんでした。\n\n"
                f"{output_pdf_path}"
            )

        return {
            "file_name_base": (
                file_name_base
            ),
            "excel_path": str(
                output_excel_path
            ),
            "pdf_path": str(
                output_pdf_path
            ),
        }

    # =====================================
    # 同名ファイル回避
    # =====================================
    def create_unique_file_name_base(
        self,
        output_directory: Path,
        file_name_base: str,
    ) -> str:
        """
        ExcelまたはPDFのどちらかが既に存在する場合、
        元ファイルを残したまま連番付きの名前を返す。

        例:
            第3シート設計部_DAS情報解析
            第3シート設計部_DAS情報解析_2
            第3シート設計部_DAS情報解析_3
        """

        candidate = file_name_base
        number = 2

        while (
            (
                output_directory
                / f"{candidate}.xlsx"
            ).exists()
            or (
                output_directory
                / f"{candidate}.pdf"
            ).exists()
        ):
            suffix = f"_{number}"

            max_base_length = max(
                1,
                self.MAX_FILE_NAME_LENGTH
                - len(suffix),
            )

            candidate = (
                file_name_base[:max_base_length]
                + suffix
            )

            number += 1

        return candidate

    # =====================================
    # ExcelWriter用データ準備
    # =====================================
    def prepare_data(
        self,
        data,
    ) -> None:
        """
        既存ExcelWriterで使用する属性名へ値を合わせる。

        特調以外TBの表示ルール:
        ・各項目の1行目を基本表示名にする。
        ・同じ1行目が複数ある場合だけ2行目を使って区別する。
        ・2行目追加後も7行以内なら、2行目をNo.なしの継続行にする。
        ・7行を超える場合は、重複項目だけ1行目+2行目を結合し、
          WorkItemShortenerで1行表示にまとめる。
        """

        application_no = str(
            getattr(
                data,
                "application_no",
                "",
            )
            or ""
        ).strip()

        setattr(
            data,
            "voucher_no",
            "",
        )

        setattr(
            data,
            "application_no",
            application_no,
        )

        setattr(
            data,
            "output_layout_mode",
            "tokucho_other",
        )

        subject = self.get_subject(data)

        if subject:
            setattr(
                data,
                "subject",
                subject,
            )
            setattr(
                data,
                "job_title",
                subject,
            )

        work_items = self.get_work_items(
            data
        )

        second_lines = (
            self.get_work_item_second_lines(
                data=data,
                item_count=len(work_items),
            )
        )

        title_counts: dict[str, int] = {}

        for title in work_items:
            key = str(title).strip()
            title_counts[key] = (
                title_counts.get(key, 0)
                + 1
            )

        duplicate_indexes = {
            index
            for index, title in enumerate(
                work_items
            )
            if (
                title_counts.get(
                    str(title).strip(),
                    0,
                ) > 1
                and index < len(second_lines)
                and bool(second_lines[index])
            )
        }

        required_rows = (
            len(work_items)
            + len(duplicate_indexes)
        )

        display_rows: list[dict] = []
        display_outputs: list[str] = []

        if required_rows <= 7:
            # 空き行あり:
            # 重複タイトルだけ2行目をNo.なしで追加。
            for index, title in enumerate(
                work_items
            ):
                title_text = (
                    self.work_item_shortener
                    .shorten(title)
                )

                display_rows.append(
                    {
                        "number": index + 1,
                        "text": title_text,
                    }
                )
                display_outputs.append(
                    title_text
                )

                if index in duplicate_indexes:
                    detail_text = (
                        self.work_item_shortener
                        .shorten(
                            second_lines[index]
                        )
                    )

                    if detail_text:
                        display_rows.append(
                            {
                                "number": None,
                                "text": detail_text,
                            }
                        )
                        display_outputs.append(
                            detail_text
                        )

        else:
            # 7行を超える:
            # 重複タイトルだけ2行目と結合して
            # 区別できる1行名称にする。
            for index, title in enumerate(
                work_items
            ):
                source_text = str(
                    title
                ).strip()

                if index in duplicate_indexes:
                    source_text = (
                        f"{source_text} "
                        f"{second_lines[index]}"
                    ).strip()

                display_text = (
                    self.work_item_shortener
                    .shorten(
                        source_text
                    )
                )

                display_rows.append(
                    {
                        "number": index + 1,
                        "text": display_text,
                    }
                )
                display_outputs.append(
                    display_text
                )

        setattr(
            data,
            "display_output_rows",
            display_rows,
        )

        setattr(
            data,
            "display_outputs",
            display_outputs,
        )

        setattr(
            data,
            "items",
            display_outputs,
        )

        setattr(
            data,
            "deliverables",
            display_outputs,
        )

        due_date = self.get_due_date(
            data
        )

        if due_date:
            setattr(
                data,
                "due_date",
                due_date,
            )
            setattr(
                data,
                "deadline",
                due_date,
            )

    def get_work_item_second_lines(
        self,
        data,
        item_count: int,
    ) -> list[str]:
        """
        Readerが保持した各項目の2行目を、
        output_titlesと同じ件数・順序で返す。

        空欄も位置合わせのため保持する。
        """

        values = getattr(
            data,
            "output_second_lines",
            None,
        )

        if values is None:
            values = []

        if isinstance(values, str):
            values = values.splitlines()

        result = [
            str(value).strip()
            if value is not None
            else ""
            for value in values
        ]

        if len(result) < item_count:
            result.extend(
                [
                    ""
                    for _ in range(
                        item_count
                        - len(result)
                    )
                ]
            )

        return result[
            :item_count
        ]

    # =====================================
    # 出力ファイル名生成
    # =====================================
    def create_output_file_name(
        self,
        department: str,
        work_item: str,
    ) -> str:
        """
        「部署名_作業項目」形式の
        ファイル名を生成する。

        Windowsで使用できない文字は
        全角またはアンダースコアへ置換する。
        """

        safe_department = (
            self.sanitize_file_name_part(
                department
            )
        )

        safe_work_item = (
            self.sanitize_file_name_part(
                work_item
            )
        )

        if not safe_department:
            safe_department = "部署名不明"

        if not safe_work_item:
            safe_work_item = "作業項目不明"

        file_name = (
            f"{safe_department}_"
            f"{safe_work_item}"
        )

        file_name = file_name.strip(
            " ._"
        )

        if len(
            file_name
        ) > self.MAX_FILE_NAME_LENGTH:
            file_name = file_name[
                :self.MAX_FILE_NAME_LENGTH
            ].rstrip(
                " ._"
            )

        if not file_name:
            raise ValueError(
                "出力ファイル名を"
                "生成できませんでした。"
            )

        return file_name

    # =====================================
    # ファイル名文字列整形
    # =====================================
    def sanitize_file_name_part(
        self,
        value,
    ) -> str:
        """
        Windowsのファイル名に使用できない文字を整形する。

        例:
            第3設計部/開発
                ↓
            第3設計部_開発
        """

        if value is None:
            return ""

        text = str(
            value
        ).strip()

        # 改行とタブを空白へ
        text = re.sub(
            r"[\r\n\t]+",
            " ",
            text,
        )

        # Windows禁止文字
        text = re.sub(
            r'[<>:"/\\|?*]',
            "_",
            text,
        )

        # 制御文字を削除
        text = re.sub(
            r"[\x00-\x1f]",
            "",
            text,
        )

        # 連続空白を1つへ
        text = re.sub(
            r"\s+",
            " ",
            text,
        )

        # 連続アンダースコアを1つへ
        text = re.sub(
            r"_+",
            "_",
            text,
        )

        return text.strip(
            " ._"
        )

    # =====================================
    # 部署名取得
    # =====================================
    def get_department(
        self,
        data,
    ) -> str:

        return str(
            getattr(
                data,
                "department",
                "",
            )
            or ""
        ).strip()

    # =====================================
    # 件名取得
    # =====================================
    def get_subject(
        self,
        data,
    ) -> str:

        return str(
            getattr(
                data,
                "subject",
                "",
            )
            or getattr(
                data,
                "job_title",
                "",
            )
            or ""
        ).strip()

    # =====================================
    # 作業項目一覧取得
    # =====================================
    def get_work_items(
        self,
        data,
    ) -> list[str]:
        """
        見積書へ表示する作業項目一覧を取得する。

        特調以外TBではReaderが生成した
        output_titles（各作業項目の1行目）を最優先する。

        output_titlesが存在しない既存データでは、
        outputs / items / deliverables の順でフォールバックする。
        """

        values = (
            getattr(
                data,
                "output_titles",
                None,
            )
            or getattr(
                data,
                "outputs",
                None,
            )
            or getattr(
                data,
                "items",
                None,
            )
            or getattr(
                data,
                "deliverables",
                None,
            )
            or []
        )

        if isinstance(
            values,
            str,
        ):
            values = [
                line.strip()
                for line in values.splitlines()
                if line.strip()
            ]

        # 同じ見出しが別項目として複数存在する場合があるため、
        # 重複排除は行わず、元の順番と件数を維持する。
        return [
            str(value).strip()
            for value in values
            if str(value).strip()
        ]

    # =====================================
    # 先頭作業項目取得
    # =====================================
    def get_first_work_item(
        self,
        data,
    ) -> str:
        """
        作業内容の先頭項目を取得する。

        現在のファイル名生成では件名を使用しているため、
        主に互換用として残している。
        """

        work_items = self.get_work_items(
            data
        )

        if work_items:
            return work_items[0]

        return self.get_subject(
            data
        )

    # =====================================
    # 納期取得
    # =====================================
    def get_due_date(
        self,
        data,
    ) -> str:

        return str(
            getattr(
                data,
                "due_date",
                "",
            )
            or getattr(
                data,
                "deadline",
                "",
            )
            or ""
        ).strip()