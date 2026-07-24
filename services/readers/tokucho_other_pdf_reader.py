import re
from pathlib import Path

import pdfplumber

from models.estimate_data import EstimateData


class TokuchoOtherPdfReader:
    """
    特調TB以外の業務委託計画書PDFを読み取る。

    表形式の帳票を優先して解析し、表から取得できない項目は
    PDF全文テキストから補完する。

    抽出項目:
        ・計画部署
        ・件名
        ・業務委託計画書No（タイトル横にある場合のみ）
        ・（参考）作業内容
        ・委託金額
        ・納期
        ・車種コード
    """

    CIRCLED_NUMBERS = "①②③④⑤⑥⑦⑧⑨⑩"

    def parse(self, pdf_path: str) -> EstimateData:
        """
        PDFを読み取り、EstimateDataへ格納して返す。
        """

        path = Path(pdf_path)

        if not path.exists():
            raise FileNotFoundError(
                "PDFファイルが見つかりません。\n\n"
                f"{path}"
            )

        if path.suffix.lower() != ".pdf":
            raise ValueError(
                "PDFファイルを指定してください。\n\n"
                f"{path.name}"
            )

        text, tables = self.extract_document(path)

        if not text.strip() and not tables:
            raise ValueError(
                "PDFから文字情報を取得できませんでした。\n\n"
                f"{path.name}"
            )

        main_table = self.select_main_table(tables)

        department = (
            self.extract_department_from_table(main_table)
            or self.extract_department_from_text(text)
        )

        subject = (
            self.extract_subject_from_table(main_table)
            or self.extract_subject_from_text(text)
        )

        application_no = (
            self.extract_application_no_from_tables(tables)
            or self.extract_application_no_from_text(text)
        )

        outputs = (
            self.extract_work_items_from_table(main_table)
            or self.extract_work_items_from_text(text)
        )

        amount = (
            self.extract_amount_from_table(main_table)
            or self.extract_amount_from_text(text)
        )

        due_date = (
            self.extract_due_date_from_table(main_table)
            or self.extract_due_date_from_text(text)
        )

        model_code = (
            self.extract_model_code_from_table(main_table)
            or self.extract_model_code_from_text(text)
        )

        missing_fields = []

        if not department:
            missing_fields.append("計画部署")

        if not subject:
            missing_fields.append("件名")

        if not outputs:
            missing_fields.append("（参考）作業内容")

        if not amount:
            missing_fields.append("委託金額")

        if not due_date:
            missing_fields.append("納期")

        if missing_fields:
            raise ValueError(
                "PDFから次の項目を取得できませんでした。\n\n"
                + "\n".join(
                    f"・{field_name}"
                    for field_name in missing_fields
                )
                + "\n\n"
                f"対象PDF：{path.name}"
            )

        data = EstimateData()
        data.department = department
        data.subject = subject
        data.application_no = application_no
        data.outputs = outputs
        data.amount = amount
        data.due_date = due_date
        data.model_code = model_code

        return data

    # =====================================
    # PDF全文・表取得
    # =====================================
    def extract_document(
        self,
        pdf_path: Path,
    ) -> tuple[str, list[list[list[str | None]]]]:
        """
        全ページのテキストと表を取得する。
        """

        text_parts: list[str] = []
        tables: list[list[list[str | None]]] = []

        try:
            with pdfplumber.open(str(pdf_path)) as pdf:
                for page in pdf.pages:
                    page_text = page.extract_text() or ""

                    if page_text.strip():
                        text_parts.append(page_text)

                    page_tables = page.extract_tables() or []

                    for table in page_tables:
                        if table:
                            tables.append(table)

        except Exception as error:
            raise RuntimeError(
                "PDFの読み取りに失敗しました。\n\n"
                f"{pdf_path.name}\n\n"
                f"{error}"
            ) from error

        return "\n".join(text_parts), tables

    def select_main_table(
        self,
        tables: list[list[list[str | None]]],
    ) -> list[list[str | None]]:
        """
        「計画部署」「委託金額」などを含む主表を選ぶ。
        """

        if not tables:
            return []

        best_table: list[list[str | None]] = []
        best_score = -1

        keywords = (
            "計画部署",
            "件名",
            "OUTPUT",
            "委託金額",
            "納期",
            "りん議",
        )

        for table in tables:
            flattened = " ".join(
                self.clean_cell(cell)
                for row in table
                for cell in row
                if cell is not None
            )

            score = sum(
                10
                for keyword in keywords
                if keyword in flattened
            ) + len(table)

            if score > best_score:
                best_score = score
                best_table = table

        return best_table

    # =====================================
    # 計画部署
    # =====================================
    def extract_department_from_table(
        self,
        table: list[list[str | None]],
    ) -> str:
        """
        計画部署行から「○○部」を取得する。
        室名はファイル名・台帳検索の対象には含めない。
        """

        for row in table:
            if not self.row_contains(row, "計画部署"):
                continue

            label_index = self.find_cell_index(row, "計画部署")

            if label_index is None:
                continue

            values = [
                self.clean_cell(value)
                for value in row[label_index + 1:]
                if self.clean_cell(value)
            ]

            department_name = ""

            for value in values:
                if value in ("部", "室", "連絡先"):
                    continue

                if self.looks_like_phone_number(value):
                    continue

                department_name = value
                break

            if not department_name:
                continue

            if not department_name.endswith("部"):
                department_name += "部"

            return department_name

        return ""

    def extract_department_from_text(self, text: str) -> str:
        patterns = [
            (
                r"計画部署\s*\n?"
                r"([^\n]+?)\s+部(?:\s|$)"
            ),
            (
                r"計画部署\s*[:：]?\s*"
                r"([^\n]+?部)"
            ),
        ]

        for pattern in patterns:
            match = re.search(pattern, text)

            if match:
                value = self.clean_single_line(match.group(1))

                if value and not value.endswith("部"):
                    value += "部"

                return value

        return ""

    # =====================================
    # 件名
    # =====================================
    def extract_subject_from_table(
        self,
        table: list[list[str | None]],
    ) -> str:
        """
        「車種」「件名」の次行にある件名を取得する。
        """

        for row_index, row in enumerate(table):
            subject_index = self.find_cell_index(row, "件名")

            if subject_index is None:
                continue

            # 同じ行の右側に値がある帳票
            same_row_value = self.first_meaningful_value(
                row[subject_index + 1:]
            )

            if (
                same_row_value
                and same_row_value != "部位コード"
            ):
                return same_row_value

            # 今回の帳票では次行の同じ列に件名が入る
            if row_index + 1 < len(table):
                next_row = table[row_index + 1]

                if subject_index < len(next_row):
                    value = self.clean_cell(
                        next_row[subject_index]
                    )

                    if value:
                        return value

        return ""

    def extract_subject_from_text(self, text: str) -> str:
        patterns = [
            r"車種\s+件名\s*\n[^\n]+\s+(.+)",
            r"業務委託計画書\s*\n(.+?)\s*\n件名",
            r"件名\s*[:：]?\s*(.+)",
        ]

        for pattern in patterns:
            match = re.search(pattern, text)

            if match:
                value = self.clean_single_line(match.group(1))

                if value and value != "連絡先":
                    return value

        return ""

    # =====================================
    # 業務委託計画書No（任意）
    # =====================================
    def extract_application_no_from_tables(
        self,
        tables: list[list[list[str | None]]],
    ) -> str:
        """
        タイトル「業務委託計画書」の横、または同じセル内に
        記載された番号だけを取得する。

        番号がない帳票では空文字を返す。
        りん議No.は業務委託計画書Noとして使用しない。
        """

        for table in tables:
            for row in table:
                for index, cell in enumerate(row):
                    value = self.clean_cell(cell)

                    if "業務委託計画書" not in value:
                        continue

                    # 同じセル内:
                    # 業務委託計画書 No.12345
                    # 業務委託計画書 ITK12345
                    same_cell = self.extract_title_number(value)

                    if same_cell:
                        return same_cell

                    # タイトルの右隣のセル
                    for next_cell in row[index + 1:]:
                        candidate = self.clean_cell(next_cell)

                        if not candidate:
                            continue

                        number = self.extract_number_candidate(candidate)

                        if number:
                            return number

                        # 最初の値が番号でなければ、その行には番号なし
                        break

        return ""

    def extract_application_no_from_text(
        self,
        text: str,
    ) -> str:
        """
        全文テキストから、タイトルと同じ行にある番号だけを取得する。
        次の行や、りん議No.は取得しない。
        """

        for line in text.splitlines():
            if "業務委託計画書" not in line:
                continue

            number = self.extract_title_number(line)

            if number:
                return number

        return ""

    def extract_title_number(
        self,
        value: str,
    ) -> str:
        """
        「業務委託計画書」より後ろの部分から番号を取得する。
        """

        title_position = value.find("業務委託計画書")

        if title_position < 0:
            return ""

        suffix = value[
            title_position + len("業務委託計画書"):
        ].strip()

        suffix = re.sub(
            r"^(?:No\.?|NO\.?|Ｎｏ\.?|番号)\s*[:：]?\s*",
            "",
            suffix,
            flags=re.IGNORECASE,
        )

        return self.extract_number_candidate(suffix)

    def extract_number_candidate(
        self,
        value: str,
    ) -> str:
        """
        業務委託計画書Noとして使える番号候補を取得する。
        英数字を含み、数字が最低1文字ある値だけを許可する。
        """

        normalized = self.normalize_ascii(
            self.clean_single_line(value)
        )

        if not normalized:
            return ""

        match = re.fullmatch(
            r"[A-Za-z0-9][A-Za-z0-9_-]*",
            normalized,
        )

        if not match:
            return ""

        if not any(character.isdigit() for character in normalized):
            return ""

        return normalized

    # =====================================
    # （参考）作業内容
    # =====================================
    def extract_work_items_from_table(
        self,
        table: list[list[str | None]],
    ) -> list[str]:
        """
        「（参考）作業内容」欄の「項目」を取得する。

        「合計」行、予定工数、空行は対象外。
        """

        items: list[str] = []
        section_started = False
        item_column_index: int | None = None

        for row in table:
            row_text = " ".join(
                self.clean_cell(cell)
                for cell in row
                if self.clean_cell(cell)
            )

            if (
                "作業内容" in row_text
                and "予定工数" in row_text
            ):
                section_started = True

                for index, cell in enumerate(row):
                    cleaned = self.clean_cell(cell)

                    if cleaned in ("項目", "項 目"):
                        item_column_index = index
                        break

                continue

            if not section_started:
                continue

            # 委託先以降は別セクション
            if self.row_contains(row, "委託先"):
                break

            value = ""

            if (
                item_column_index is not None
                and item_column_index < len(row)
            ):
                value = self.clean_cell(
                    row[item_column_index]
                )

            # 表によっては「項目」の位置がずれるため、
            # 合計・工数以外の最初の有効セルを補助的に使用
            if not value:
                for cell in row:
                    candidate = self.clean_cell(cell)

                    if not candidate:
                        continue

                    if re.fullmatch(
                        r"\d+(?:\.\d+)?\s*Hr",
                        candidate,
                        flags=re.IGNORECASE,
                    ):
                        continue

                    value = candidate
                    break

            value = self.clean_work_item(value)

            if not value:
                continue

            if value in (
                "合計",
                "項目",
                "項 目",
                "（参考）作業内容",
            ):
                continue

            if "予定工数" in value:
                continue

            items.append(value)

        return self.unique_nonempty(items)

    def extract_work_items_from_text(
        self,
        text: str,
    ) -> list[str]:
        """
        表抽出に失敗した場合の補助処理。
        「作業内容」から「合計」までを対象にする。
        """

        match = re.search(
            r"(?:（参考）|\(参考\))?\s*作業内容"
            r".*?(?:項\s*目)"
            r"(.*?)"
            r"(?=\n\s*合計|\n\s*委託先)",
            text,
            flags=re.DOTALL,
        )

        if not match:
            return []

        items: list[str] = []

        for line in match.group(1).splitlines():
            cleaned = self.clean_single_line(line)

            if not cleaned:
                continue

            cleaned = re.sub(
                r"\s+\d+(?:\.\d+)?\s*Hr\s*$",
                "",
                cleaned,
                flags=re.IGNORECASE,
            ).strip()

            cleaned = self.clean_work_item(cleaned)

            if not cleaned or cleaned == "合計":
                continue

            items.append(cleaned)

        return self.unique_nonempty(items)

    # =====================================
    # 委託金額
    # =====================================
    def extract_amount_from_table(
        self,
        table: list[list[str | None]],
    ) -> str:
        for row in table:
            if not self.row_contains(row, "委託金額"):
                continue

            row_text = " ".join(
                self.clean_cell(cell)
                for cell in row
                if self.clean_cell(cell)
            )

            amount = self.find_amount(row_text)

            if amount:
                return amount

        return ""

    def extract_amount_from_text(self, text: str) -> str:
        match = re.search(
            r"委託金額\s*"
            r"([￥¥]?\s*[\d,，]+(?:\.\d+)?\s*円?)",
            text,
        )

        if not match:
            return ""

        return self.normalize_amount_text(match.group(1))

    # =====================================
    # 納期
    # =====================================
    def extract_due_date_from_table(
        self,
        table: list[list[str | None]],
    ) -> str:
        """
        表のセル結合により日付が委託金額セルに入る場合があるため、
        主表全体から日付を集め、開始予定日の次の日付を納期とする。
        """

        all_text = "\n".join(
            " ".join(
                self.clean_cell(cell)
                for cell in row
                if self.clean_cell(cell)
            )
            for row in table
        )

        return self.extract_due_date_from_text(all_text)

    def extract_due_date_from_text(self, text: str) -> str:
        direct_match = re.search(
            r"納期\s*"
            r"(\d{4})年\s*"
            r"(\d{1,2})月\s*"
            r"(\d{1,2})日",
            text,
        )

        if direct_match:
            return self.format_japanese_date(
                direct_match.group(1),
                direct_match.group(2),
                direct_match.group(3),
            )

        # pdfplumberの表抽出で「納期」と日付が別セルになる場合
        dates = re.findall(
            r"(\d{4})年\s*"
            r"(\d{1,2})月\s*"
            r"(\d{1,2})日",
            text,
        )

        if len(dates) >= 2:
            year, month, day = dates[1]

            return self.format_japanese_date(
                year,
                month,
                day,
            )

        return ""

    # =====================================
    # 車種コード
    # =====================================
    def extract_model_code_from_table(
        self,
        table: list[list[str | None]],
    ) -> str:
        for row_index, row in enumerate(table):
            model_index = self.find_cell_index(row, "車種")

            if model_index is None:
                continue

            same_row_value = self.first_meaningful_value(
                row[model_index + 1:]
            )

            if (
                same_row_value
                and same_row_value != "件名"
            ):
                return same_row_value

            if row_index + 1 < len(table):
                next_row = table[row_index + 1]

                if model_index < len(next_row):
                    value = self.clean_cell(
                        next_row[model_index]
                    )

                    if value:
                        return value

        return ""

    def extract_model_code_from_text(self, text: str) -> str:
        patterns = [
            r"車種\s+件名\s*\n([A-Za-z0-9_-]+)",
            r"車種\s*\n([A-Za-z0-9_-]+)",
            r"車種\s*[:：]?\s*([A-Za-z0-9_-]+)",
        ]

        for pattern in patterns:
            match = re.search(
                pattern,
                text,
                flags=re.IGNORECASE,
            )

            if match:
                return match.group(1).strip()

        return ""

    # =====================================
    # 共通補助処理
    # =====================================
    def find_amount(self, text: str) -> str:
        match = re.search(
            r"([￥¥]?\s*[\d,，]+(?:\.\d+)?\s*円)",
            text,
        )

        if not match:
            return ""

        return self.normalize_amount_text(match.group(1))

    def normalize_amount_text(self, value: str) -> str:
        return (
            self.clean_single_line(value)
            .replace("，", ",")
            .replace(" ", "")
        )

    def format_japanese_date(
        self,
        year: str,
        month: str,
        day: str,
    ) -> str:
        return (
            f"{int(year)}年"
            f"{int(month)}月"
            f"{int(day)}日"
        )

    def clean_cell(self, value) -> str:
        if value is None:
            return ""

        text = str(value).replace("\r", "\n")
        text = re.sub(r"\s+", " ", text)

        return text.strip()

    def clean_single_line(self, value: str) -> str:
        if not value:
            return ""

        return re.sub(
            r"\s+",
            " ",
            str(value),
        ).strip()

    def clean_work_item(self, value: str) -> str:
        if not value:
            return ""

        cleaned = self.clean_single_line(value)

        cleaned = re.sub(
            r"^[①②③④⑤⑥⑦⑧⑨⑩]\s*",
            "",
            cleaned,
        )

        cleaned = re.sub(
            r"^\(?\d+\)?[\.．、\s]+",
            "",
            cleaned,
        )

        cleaned = re.sub(
            r"^[・●○■□◆◇\-－]\s*",
            "",
            cleaned,
        )

        return cleaned.strip()

    def row_contains(
        self,
        row: list[str | None],
        keyword: str,
    ) -> bool:
        return any(
            keyword in self.clean_cell(cell)
            for cell in row
        )

    def find_cell_index(
        self,
        row: list[str | None],
        keyword: str,
    ) -> int | None:
        for index, cell in enumerate(row):
            if keyword in self.clean_cell(cell):
                return index

        return None

    def find_cell_index_regex(
        self,
        row: list[str | None],
        pattern: str,
    ) -> int | None:
        for index, cell in enumerate(row):
            if re.search(
                pattern,
                self.clean_cell(cell),
                flags=re.IGNORECASE,
            ):
                return index

        return None

    def first_meaningful_value(
        self,
        values,
    ) -> str:
        for value in values:
            cleaned = self.clean_cell(value)

            if cleaned:
                return cleaned

        return ""

    def looks_like_phone_number(self, value: str) -> bool:
        return bool(
            re.fullmatch(
                r"\d{2,4}-\d{2,4}",
                value,
            )
        )

    def normalize_ascii(self, value: str) -> str:
        translation = str.maketrans(
            "ＡＢＣＤＥＦＧＨＩＪＫＬＭＮＯＰＱＲＳＴＵＶＷＸＹＺ"
            "ａｂｃｄｅｆｇｈｉｊｋｌｍｎｏｐｑｒｓｔｕｖｗｘｙｚ"
            "０１２３４５６７８９",
            "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
            "abcdefghijklmnopqrstuvwxyz"
            "0123456789",
        )

        return value.translate(translation).strip()

    def unique_nonempty(
        self,
        values: list[str],
    ) -> list[str]:
        result: list[str] = []
        seen: set[str] = set()

        for value in values:
            cleaned = self.clean_single_line(value)

            if not cleaned or cleaned in seen:
                continue

            seen.add(cleaned)
            result.append(cleaned)

        return result
