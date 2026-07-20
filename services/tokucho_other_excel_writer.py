import re
from pathlib import Path

from services.template_resolver import TemplateResolver
from services.excel_writer import ExcelWriter


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
                work_item=self.get_first_work_item(
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

        特調TB以外では、
        業務委託計画書横の番号を
        従来の伝票番号と同じC25へ転記するため、
        voucher_noへ設定する。
        """

        application_no = str(
            getattr(
                data,
                "application_no",
                "",
            )
            or ""
        ).strip()

        # 業務委託計画書Noは帳票によって存在しない。
        # 値がない場合はvoucher_noを空にし、C25へ転記しない。
        setattr(
            data,
            "voucher_no",
            application_no,
        )

        # ExcelWriterがjob_titleを使う場合に備える
        subject = self.get_subject(
            data
        )

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

        # ExcelWriterがitemsまたはdeliverablesを
        # 使用する場合に備えて両方へ設定
        work_items = self.get_work_items(
            data
        )

        setattr(
            data,
            "items",
            work_items,
        )

        setattr(
            data,
            "deliverables",
            work_items,
        )

        # ExcelWriterがdue_dateまたはdeadlineを
        # 使用する場合に備えて両方へ設定
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

        values = (
            getattr(
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
        ファイル名には、作業内容の
        先頭項目を使用する。
        """

        work_items = self.get_work_items(
            data
        )

        if work_items:
            return work_items[0]

        # 作業項目が取得できない場合は、
        # 件名を代替として使用する
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