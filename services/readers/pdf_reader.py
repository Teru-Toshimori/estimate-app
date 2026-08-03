import logging
import re
from pathlib import Path

import pdfplumber

from models.estimate_data import EstimateData
from services.ai.tokucho_pdf_openai_vision import (
    extract_tokucho_pdf_data,
)


logger = logging.getLogger(__name__)


class PDFReader:
    """
    特調TB用の業務委託計画書を読み取る。

    通常PDF：
        pdfplumberで文字列を取得して解析する。

    画像PDF：
        pdfplumberで十分な文字が取得できない場合、
        OpenAI Visionへ自動的に切り替える。
    """

    # PDFからこの文字数以上取得できた場合は、
    # 通常のテキストPDFとして処理する。
    MIN_TEXT_LENGTH = 50

    def read_text(
        self,
        pdf_path: str | Path,
    ) -> str:
        """
        PDF全ページから文字列を取得する。
        """

        path = Path(pdf_path)
        text_parts: list[str] = []

        if not path.exists():
            raise FileNotFoundError(
                f"PDFファイルが見つかりません。\n{path}"
            )

        try:
            with pdfplumber.open(
                str(path)
            ) as pdf:
                for page in pdf.pages:
                    page_text = (
                        page.extract_text()
                        or ""
                    )

                    if page_text.strip():
                        text_parts.append(
                            page_text
                        )

        except Exception as error:
            logger.exception(
                "PDF読み取り失敗: %s",
                path,
            )

            raise RuntimeError(
                "PDFの読み取りに失敗しました。\n\n"
                f"対象PDF：{path.name}\n"
                f"詳細：{error}"
            ) from error

        return "\n".join(
            text_parts
        )

    def parse(
        self,
        pdf_path: str | Path,
    ) -> EstimateData:
        """
        PDFを解析してEstimateDataを返す。
        """

        path = Path(pdf_path)

        text = self.read_text(
            path
        )

        # ---------------------------------
        # 画像PDF判定
        # ---------------------------------
        if self.is_image_pdf(text):
            logger.info(
                "画像PDFを検出しました。AI解析へ切り替えます: %s",
                path.name,
            )

            return self.parse_with_openai(
                path
            )

        # ---------------------------------
        # 通常の文字PDF
        # ---------------------------------
        logger.info(
            "文字PDFとして解析します: %s",
            path.name,
        )

        return self.parse_text_pdf(
            text=text,
            pdf_path=path,
        )

    def is_image_pdf(
        self,
        text: str,
    ) -> bool:
        """
        pdfplumberで十分な文字が取得できなければ
        画像PDFと判断する。
        """

        normalized_text = re.sub(
            r"\s+",
            "",
            text or "",
        )

        return (
            len(normalized_text)
            < self.MIN_TEXT_LENGTH
        )

    def parse_with_openai(
        self,
        pdf_path: str | Path,
    ) -> EstimateData:
        """
        画像PDFをOpenAI Visionで解析する。
        """

        result = extract_tokucho_pdf_data(
            pdf_path
        )

        data = EstimateData()

        data.application_no = result.get(
            "application_no",
            "",
        )

        data.department = self.normalize_department(
            result.get(
                "department",
                "",
            )
        )

        data.subject = result.get(
            "subject",
            "",
        )

        data.model_code = result.get(
            "model_code",
            "",
        )

        data.amount = result.get(
            "amount",
            "",
        )

        data.due_date = result.get(
            "due_date",
            "",
        )

        data.outputs = result.get(
            "outputs",
            [],
        )

        self.validate_data(
            data=data,
            pdf_path=Path(pdf_path),
        )

        return data

    def parse_text_pdf(
        self,
        text: str,
        pdf_path: Path,
    ) -> EstimateData:
        """
        文字情報を持つPDFを従来方式で解析する。
        """

        lines = text.splitlines()

        data = EstimateData()

        # 基本情報
        data.application_no = self.get_application_no(
            lines
        )

        data.department = self.normalize_department(
            self.get_department(
                lines
            )
        )

        data.subject = self.get_subject(
            lines
        )

        # 車種コード
        data.model_code = self.get_model_code(
            lines
        )

        # 金額・納期
        data.amount = self.get_amount(
            lines
        )

        data.due_date = self.get_due_date(
            lines
        )

        # 成果物
        data.outputs = self.get_outputs(
            lines
        )

        self.validate_data(
            data=data,
            pdf_path=pdf_path,
        )

        return data

    def validate_data(
        self,
        data: EstimateData,
        pdf_path: Path,
    ) -> None:
        """
        見積書作成に必要な項目が取得できたか確認する。
        """

        missing_fields: list[str] = []

        if not data.application_no:
            missing_fields.append(
                "申請書No"
            )

        if not data.department:
            missing_fields.append(
                "依頼部署"
            )

        if not data.subject:
            missing_fields.append(
                "件名"
            )

        if not data.outputs:
            missing_fields.append(
                "成果物"
            )

        if not data.amount:
            missing_fields.append(
                "委託金額"
            )

        if not data.due_date:
            missing_fields.append(
                "納期"
            )

        if missing_fields:
            raise ValueError(
                "PDFから次の項目を取得できませんでした。\n\n"
                + "\n".join(
                    f"・{field_name}"
                    for field_name in missing_fields
                )
                + "\n\n"
                f"対象PDF：{pdf_path.name}"
            )

    # =====================================
    # 文字列の正規化
    # =====================================

    def normalize_department(
        self,
        department: str,
    ) -> str:
        """
        部署名に含まれる半角・全角スペースを除去する。

        例：
        R rシート骨格設計部
            ↓
        Rrシート骨格設計部
        """

        return (
            department
            .replace(" ", "")
            .replace("　", "")
            .strip()
        )

    def clean_value(
        self,
        value: str,
    ) -> str:
        """
        抽出した1行の前後空白を整理する。
        """

        return re.sub(
            r"\s+",
            " ",
            value,
        ).strip()

    # =====================================
    # 申請書No
    # =====================================

    def get_application_no(
        self,
        lines: list[str],
    ) -> str:
        """
        申請書Noを取得する。
        """

        patterns = [
            r"\bITK\d+\b",
            r"申請書\s*NO\.?\s*[:：]?\s*([A-Za-z0-9-]+)",
            r"Application\s*No\.?\s*[:：]?\s*([A-Za-z0-9-]+)",
        ]

        for line in lines:
            for pattern in patterns:
                match = re.search(
                    pattern,
                    line,
                    flags=re.IGNORECASE,
                )

                if match:
                    if match.lastindex:
                        return (
                            match.group(1)
                            .strip()
                        )

                    return (
                        match.group(0)
                        .strip()
                    )

        return ""

    # =====================================
    # 依頼部署
    # =====================================

    def get_department(
        self,
        lines: list[str],
    ) -> str:
        """
        依頼部署を取得する。

        次の形式に対応する。

        依頼部署
        Request Div.
        第3シート設計部

        または

        依頼部署 Request Div. 第3シート設計部
        """

        ignored_values = {
            "依頼部署",
            "Request Div.",
            "Request Div",
            "Request Division",
        }

        for index, line in enumerate(lines):
            stripped_line = line.strip()

            if (
                "依頼部署" not in stripped_line
                and "Request Div" not in stripped_line
            ):
                continue

            # 同じ行に部署名がある場合
            same_line = re.sub(
                r"依頼部署",
                "",
                stripped_line,
            )

            same_line = re.sub(
                r"Request\s+Div(?:ision)?\.?",
                "",
                same_line,
                flags=re.IGNORECASE,
            )

            same_line = same_line.strip(
                " ：:"
            )

            if (
                same_line
                and same_line not in ignored_values
            ):
                return self.clean_value(
                    same_line
                )

            # 後続行から部署名を探す
            search_end = min(
                index + 5,
                len(lines),
            )

            for next_index in range(
                index + 1,
                search_end,
            ):
                value = lines[
                    next_index
                ].strip()

                if not value:
                    continue

                if value in ignored_values:
                    continue

                if (
                    value.startswith("Request Div")
                    or value.startswith("Request Division")
                ):
                    continue

                return self.clean_value(
                    value
                )

        return ""

    # =====================================
    # 件名
    # =====================================

    def get_subject(
        self,
        lines: list[str],
    ) -> str:
        """
        件名を取得する。
        """

        ignored_values = {
            "件名",
            "Job Title",
        }

        for index, line in enumerate(lines):
            stripped_line = line.strip()

            if (
                "件名" not in stripped_line
                and "Job Title" not in stripped_line
            ):
                continue

            same_line = re.sub(
                r"件名",
                "",
                stripped_line,
            )

            same_line = re.sub(
                r"Job\s+Title",
                "",
                same_line,
                flags=re.IGNORECASE,
            )

            same_line = same_line.strip(
                " ：:"
            )

            if (
                same_line
                and same_line not in ignored_values
            ):
                return self.clean_value(
                    same_line
                )

            search_end = min(
                index + 5,
                len(lines),
            )

            for next_index in range(
                index + 1,
                search_end,
            ):
                value = lines[
                    next_index
                ].strip()

                if not value:
                    continue

                if value in ignored_values:
                    continue

                if value.startswith(
                    "Job Title"
                ):
                    continue

                return self.clean_value(
                    value
                )

        return ""

    # =====================================
    # 車種コード
    # =====================================

    def get_model_code(
        self,
        lines: list[str],
    ) -> str:
        """
        車種コードを取得する。
        """

        for index, line in enumerate(lines):
            stripped_line = line.strip()

            if (
                "車種コード" not in stripped_line
                and "Model Code" not in stripped_line
            ):
                continue

            same_line = re.sub(
                r"車種コード",
                "",
                stripped_line,
            )

            same_line = re.sub(
                r"Model\s+Code",
                "",
                same_line,
                flags=re.IGNORECASE,
            )

            same_line = same_line.strip(
                " ：:"
            )

            if same_line:
                return (
                    same_line
                    .split()[0]
                )

            search_end = min(
                index + 4,
                len(lines),
            )

            for next_index in range(
                index + 1,
                search_end,
            ):
                value = lines[
                    next_index
                ].strip()

                if not value:
                    continue

                if value in {
                    "Model Code",
                    "車種コード",
                }:
                    continue

                return value.split()[0]

        return ""

    # =====================================
    # 委託金額
    # =====================================

    def get_amount(
        self,
        lines: list[str],
    ) -> str:
        """
        委託金額を取得する。
        """

        patterns = [
            r"委託金額\s*[:：]?\s*([\d,]+\s*円)",
            r"Outsourcing\s+Job\s+Fee\s*[:：]?\s*([\d,]+\s*円)",
            r"([\d,]+\s*円)",
        ]

        for line in lines:
            for pattern in patterns:
                match = re.search(
                    pattern,
                    line,
                    flags=re.IGNORECASE,
                )

                if match:
                    return (
                        match.group(1)
                        .replace(" ", "")
                        .strip()
                    )

        return ""

    # =====================================
    # 納期
    # =====================================

    def get_due_date(
        self,
        lines: list[str],
    ) -> str:
        """
        納期を取得する。
        """

        patterns = [
            r"納期\s*[:：]?\s*(\d{4}年\d{1,2}月\d{1,2}日)",
            (
                r"Delivery\s+Due\s+date"
                r"\s*[:：]?\s*"
                r"(\d{4}年\d{1,2}月\d{1,2}日)"
            ),
            r"納期\s*[:：]?\s*(\d{4}/\d{1,2}/\d{1,2})",
        ]

        for line in lines:
            for pattern in patterns:
                match = re.search(
                    pattern,
                    line,
                    flags=re.IGNORECASE,
                )

                if match:
                    return match.group(1)

        return ""

    # =====================================
    # 成果物
    # =====================================

    def get_outputs(
        self,
        lines: list[str],
    ) -> list[str]:
        """
        成果物一覧を取得する。
        """

        outputs: list[str] = []
        started = False

        for line in lines:
            stripped_line = line.strip()

            if (
                "成果物名称" in stripped_line
                or "Name of output" in stripped_line
            ):
                started = True
                continue

            if not started:
                continue

            if (
                "成果物に" in stripped_line
                or "Requirement" in stripped_line
                or "TBから提供する" in stripped_line
                or "Information provided" in stripped_line
            ):
                break

            match = re.match(
                r"^[①-⑳]\s*(.+)$",
                stripped_line,
            )

            if not match:
                continue

            value = match.group(1).strip()

            if value:
                outputs.append(
                    value
                )

        return outputs