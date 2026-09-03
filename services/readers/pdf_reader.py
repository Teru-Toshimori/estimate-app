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

        output_titles, output_second_lines = (
            self.build_output_display_metadata(
                data.outputs
            )
        )

        setattr(
            data,
            "output_titles",
            output_titles,
        )

        setattr(
            data,
            "output_second_lines",
            output_second_lines,
        )

        setattr(
            data,
            "output_layout_mode",
            "tokucho",
        )

        self.validate_data(
            data=data,
            pdf_path=Path(pdf_path),
            source_text="",
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

        # 見積書表示用データ
        # outputs              : 元PDFから取得した成果物全文
        # output_titles        : 各成果物の1行目
        # output_second_lines  : 各成果物の2行目
        output_titles, output_second_lines = (
            self.build_output_display_metadata(
                data.outputs
            )
        )

        setattr(
            data,
            "output_titles",
            output_titles,
        )

        setattr(
            data,
            "output_second_lines",
            output_second_lines,
        )

        setattr(
            data,
            "output_layout_mode",
            "tokucho",
        )

        self.validate_data(
            data=data,
            pdf_path=pdf_path,
            source_text=text,
        )

        return data

    def validate_data(
        self,
        data: EstimateData,
        pdf_path: Path,
        source_text: str = "",
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
                self.build_extraction_error_message(
                    pdf_path=pdf_path,
                    source_text=source_text,
                    missing_fields=missing_fields,
                )
            )

    # =====================================
    # 抽出失敗時のエラーメッセージ
    # =====================================

    def build_extraction_error_message(
        self,
        pdf_path: Path,
        source_text: str,
        missing_fields: list[str],
    ) -> str:
        """
        必須項目が複数取得できない場合、または
        特調TBの主要マーカーが不足している場合は、
        PDFフォーマット違いの可能性を案内する。
        """

        missing_text = "\n".join(
            f"・{field_name}"
            for field_name in missing_fields
        )

        if self.is_possible_format_mismatch(
            source_text=source_text,
            missing_fields=missing_fields,
        ):
            return (
                "PDFのフォーマットが対応形式と異なる可能性があります。\n\n"
                "処理種別：特調TB\n\n"
                "選択したPDFが「特調TB」の業務委託計画書"
                "フォーマットであることを確認してください。\n\n"
                "取得できなかった項目：\n"
                f"{missing_text}\n\n"
                f"対象PDF：{pdf_path.name}"
            )

        return (
            "PDFから次の項目を取得できませんでした。\n\n"
            f"{missing_text}\n\n"
            f"対象PDF：{pdf_path.name}"
        )

    def is_possible_format_mismatch(
        self,
        source_text: str,
        missing_fields: list[str],
    ) -> bool:
        """
        特調TBフォーマットと異なる可能性を簡易判定する。
        """

        if len(missing_fields) >= 2:
            return True

        combined_text = source_text or ""

        # 画像PDFではsource_textが空になるため、
        # 1項目だけの欠落をフォーマット違いとは断定しない。
        if not combined_text.strip():
            return False

        format_markers = (
            "依頼部署",
            "Request Div",
            "件名",
            "Job Title",
            "成果物名称",
            "Name of output",
            "委託金額",
            "納期",
        )

        marker_count = sum(
            1
            for marker in format_markers
            if marker in combined_text
        )

        return marker_count < 4

    # =====================================
    # 成果物表示用メタデータ
    # =====================================

    def build_output_display_metadata(
        self,
        outputs: list[str],
    ) -> tuple[list[str], list[str]]:
        """
        成果物全文から、見積書表示用の1行目・2行目を作る。

        outputsは全文を保持したまま、表示用だけを分離する。
        """

        titles: list[str] = []
        second_lines: list[str] = []

        for output in outputs or []:
            text = self._remove_output_number(
                str(output or "")
            ).strip()

            if not text:
                continue

            physical_lines = [
                self.clean_value(line)
                for line in text.splitlines()
                if self.clean_value(line)
            ]

            if not physical_lines:
                continue

            titles.append(physical_lines[0])
            second_lines.append(
                physical_lines[1]
                if len(physical_lines) >= 2
                else ""
            )

        return titles, second_lines

    def _remove_output_number(
        self,
        text: str,
    ) -> str:
        return re.sub(
            r"^[①-⑳]\s*",
            "",
            text or "",
        ).strip()

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

        特調TBのPDFでは、pdfplumberの抽出結果が次の2パターンになる。

        パターン1:
            ① 設計問題点の調査、展開
            市場不具合の確認・分類・判定根拠整理

        パターン2:
            ②
            市場不具合 発生箇所の確認
            製品図面などの発生箇所情報を確認

        ①～⑳が本文と同じ行にある場合だけでなく、
        番号だけが独立した行として抽出される場合にも対応する。

        空のテンプレート行では「④」「⑤」「⑥」のように
        番号だけが連続することがあるため、本文を持たない項目は
        成果物として追加しない。
        """

        outputs: list[str] = []
        current_lines: list[str] = []
        started = False
        waiting_for_content = False

        end_markers = (
            "成果物に",
            "Requirement",
            "TBから提供する",
            "Information provided",
        )

        def flush_current() -> None:
            nonlocal current_lines
            nonlocal waiting_for_content

            cleaned_lines = [
                self.clean_value(value)
                for value in current_lines
                if self.clean_value(value)
            ]

            if cleaned_lines:
                outputs.append(
                    "\n".join(cleaned_lines)
                )

            current_lines = []
            waiting_for_content = False

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

            if any(
                marker in stripped_line
                for marker in end_markers
            ):
                flush_current()
                break

            # 「① 本文」と「②」の両方を認識する。
            numbered_match = re.match(
                r"^[①-⑳]\s*(.*)$",
                stripped_line,
            )

            if numbered_match:
                flush_current()

                value = (
                    numbered_match
                    .group(1)
                    .strip()
                )

                if value:
                    current_lines = [value]
                    waiting_for_content = False
                else:
                    # 番号だけの行。
                    # 次の通常行をこの成果物の1行目として扱う。
                    waiting_for_content = True

                continue

            if not stripped_line:
                continue

            if waiting_for_content:
                current_lines = [
                    stripped_line
                ]
                waiting_for_content = False
                continue

            if not current_lines:
                continue

            current_lines.append(
                stripped_line
            )

        else:
            flush_current()

        return outputs

