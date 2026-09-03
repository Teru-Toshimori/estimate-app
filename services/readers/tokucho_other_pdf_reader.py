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
            self.extract_department_from_tables(tables)
            or self.extract_department_from_text(text)
        )

        model_code, subject = (
            self.extract_model_and_subject_from_tables(tables)
        )

        # 表解析で件名として見出し文字を誤取得した場合は無効化する。
        if self.is_invalid_subject(subject):
            subject = ""

        # 表解析で取得した車種コードも妥当性を確認する。
        # 例: 件名先頭の "DAS" を車種コードとして誤認した場合は無効化する。
        if self.is_invalid_model_code(model_code):
            model_code = ""

        # 件名を全文テキストから補完する。
        if not subject:
            subject = self.extract_subject_from_text(text)

        # 車種コードを全文テキストから補完する。
        if not model_code:
            model_code = self.extract_model_code_from_text(text)

        model_code = self.normalize_model_code(model_code)

        # 件名の先頭に車種コードが混ざっている場合は除去する。
        # 例: "QA23 無償修理低減解析" -> "無償修理低減解析"
        subject = self.remove_model_code_from_subject(
            subject,
            model_code,
        )

        application_no = (
            self.extract_application_no_from_tables(tables)
            or self.extract_application_no_from_text(text)
        )

        # 作業内容の全文とは別に、見積書表示用の「各項目の1行目」も保持する。
        # extract_work_items_from_table() が表から取得できた場合は、
        # 元PDFセル内の物理的な1行目を _last_work_item_titles へ保存する。
        self._last_work_item_titles: list[str] = []
        self._last_work_item_second_lines: list[str] = []

        outputs = self.extract_work_items_from_tables(tables)

        if not outputs:
            outputs = self.extract_work_items_from_text(text)

            # 全文テキスト解析へフォールバックした場合は、
            # 各取得行そのものを見積書用タイトルとして使用する。
            self._last_work_item_titles = [
                self.clean_work_item(value)
                for value in outputs
                if self.clean_work_item(value)
            ]
            self._last_work_item_second_lines = [
                ""
                for _ in self._last_work_item_titles
            ]

        # 見積書表示用タイトルは、同じ文字列でも別項目として存在し得る。
        # 例: 1番と3番がどちらも「設計問題点の調査、展開」。
        # そのため重複排除は行わず、元の項目順・件数をそのまま保持する。
        output_titles = self.nonempty_preserve_duplicates(
            getattr(self, "_last_work_item_titles", [])
        )

        # タイトル取得に失敗した場合でも従来処理を壊さない。
        if not output_titles:
            output_titles = [
                self.clean_work_item(value)
                for value in outputs
                if self.clean_work_item(value)
            ]

        # 同名タイトルを区別する際に使用する各項目の2行目。
        # 空欄も項目位置合わせのため保持する。
        raw_second_lines = list(
            getattr(
                self,
                "_last_work_item_second_lines",
                [],
            )
            or []
        )

        output_second_lines = [
            self.clean_work_item(value)
            if value else ""
            for value in raw_second_lines
        ]

        if len(output_second_lines) < len(output_titles):
            output_second_lines.extend(
                [
                    ""
                    for _ in range(
                        len(output_titles)
                        - len(output_second_lines)
                    )
                ]
            )
        elif len(output_second_lines) > len(output_titles):
            output_second_lines = output_second_lines[
                :len(output_titles)
            ]

        amount = (
            self.extract_amount_from_table(main_table)
            or self.extract_amount_from_text(text)
        )

        due_date = (
            self.extract_due_date_from_table(main_table)
            or self.extract_due_date_from_text(text)
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
                self.build_extraction_error_message(
                    path=path,
                    text=text,
                    tables=tables,
                    missing_fields=missing_fields,
                )
            )

        data = EstimateData()
        data.department = department
        data.subject = subject
        data.application_no = application_no
        # outputs      : 元PDFから取得した作業内容全文
        # output_titles: 見積書へ表示する各作業項目の1行目
        data.outputs = outputs
        setattr(data, "output_titles", output_titles)
        setattr(
            data,
            "output_second_lines",
            output_second_lines,
        )
        data.amount = amount
        data.due_date = due_date
        data.model_code = model_code

        return data

    # =====================================
    # 抽出失敗時のエラーメッセージ
    # =====================================
    def build_extraction_error_message(
        self,
        path: Path,
        text: str,
        tables: list[list[list[str | None]]],
        missing_fields: list[str],
    ) -> str:
        """
        必須項目の抽出に失敗した場合のメッセージを生成する。

        想定フォーマットの特徴が弱い場合、または複数の必須項目を
        同時に取得できなかった場合は、単なる項目欠落ではなく
        「特調以外TBで対応しているPDF形式と異なる可能性」を案内する。

        1項目だけの抽出失敗で、帳票の主要見出しが十分確認できる場合は、
        従来どおり取得できなかった項目を中心に案内する。
        """

        missing_text = "\n".join(
            f"・{field_name}"
            for field_name in missing_fields
        )

        if self.is_possible_format_mismatch(
            text=text,
            tables=tables,
            missing_fields=missing_fields,
        ):
            return (
                "PDFのフォーマットが対応形式と異なる可能性があります。\n\n"
                "処理種別：特調以外TB\n\n"
                "選択したPDFが「特調以外TB」の業務委託計画書"
                "フォーマットであることを確認してください。\n\n"
                "取得できなかった項目：\n"
                f"{missing_text}\n\n"
                f"対象PDF：{path.name}"
            )

        return (
            "PDFから次の項目を取得できませんでした。\n\n"
            f"{missing_text}\n\n"
            f"対象PDF：{path.name}"
        )

    def is_possible_format_mismatch(
        self,
        text: str,
        tables: list[list[list[str | None]]],
        missing_fields: list[str],
    ) -> bool:
        """
        特調以外TBの想定フォーマットと異なる可能性が高いか判定する。

        判定は安全側に倒し、次のどちらかを満たした場合に警告する。

        ・必須項目を2つ以上同時に取得できない
        ・主要な帳票見出しが十分に確認できない

        エラーメッセージでは断定せず「可能性があります」と表現する。
        """

        if len(missing_fields) >= 2:
            return True

        combined_text_parts = [text or ""]

        for table in tables or []:
            for row in table or []:
                combined_text_parts.extend(
                    self.clean_cell(cell)
                    for cell in row or []
                    if cell is not None
                )

        combined_text = " ".join(combined_text_parts)

        format_markers = (
            "業務委託計画書",
            "計画部署",
            "件名",
            "OUTPUT",
            "作業内容",
            "委託金額",
            "納期",
        )

        marker_count = sum(
            1
            for marker in format_markers
            if marker in combined_text
        )

        # 正常な特調以外TB帳票では主要見出しの大半が確認できる。
        # 4個未満の場合は別フォーマットの可能性が高い。
        return marker_count < 4

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
    def extract_department_from_tables(
        self,
        tables: list[list[list[str | None]]],
    ) -> str:
        """
        PDF内の全テーブルから計画部署を検索する。

        帳票によって「計画部署」が主表とは別テーブルに
        分割される場合があるため、全テーブルを対象とする。
        """

        for table in tables:
            department = self.extract_department_from_table(
                table
            )

            if department:
                return department

        return ""

    def extract_department_from_table(
        self,
        table: list[list[str | None]],
    ) -> str:
        """
        表から計画部署を取得する。

        次のような帳票差異に対応する。

        1. 「計画部署」と部署名が別セルの場合
           ['計画部署', '設計品質改善部', '連絡先']

        2. PDFの表抽出により1セルへ結合された場合
           '設計品質改善 部 連絡先 計画部署 ...'

        室名、連絡先、承認欄、予算欄などは部署名として扱わない。
        """

        ignored_values = {
            "部",
            "室",
            "連絡先",
            "承認",
            "審査",
            "作成",
            "予算",
            "車種",
            "件名",
            "部位コード",
        }

        for row in table:
            # -------------------------------------
            # パターン1：計画部署と部署名が別セル
            # -------------------------------------
            label_index = self.find_cell_index(
                row,
                "計画部署",
            )

            if label_index is not None:
                for cell in row[label_index + 1:]:
                    value = self.clean_cell(cell)

                    if not value:
                        continue

                    if value in ignored_values:
                        continue

                    if self.looks_like_phone_number(value):
                        continue

                    # 「○○部」の形ならそのまま返す。
                    if value.endswith("部"):
                        return value

                    # 「○○ 部」のように分離されている場合。
                    department_match = re.search(
                        r"([^\s]+)\s*部(?:\s|$)",
                        value,
                    )

                    if department_match:
                        return (
                            self.clean_single_line(
                                department_match.group(1)
                            )
                            + "部"
                        )

                    # 右隣セルに部署名だけが入る帳票。
                    # 「室」などの明確な除外値でなければ部署名として扱う。
                    if not any(
                        keyword in value
                        for keyword in (
                            "連絡先",
                            "承認",
                            "審査",
                            "作成",
                            "予算",
                        )
                    ):
                        return (
                            value
                            if value.endswith("部")
                            else value + "部"
                        )

            # -------------------------------------
            # パターン2：1セルへ結合された帳票
            # -------------------------------------
            for cell in row:
                raw_value = "" if cell is None else str(cell)

                if "計画部署" not in raw_value:
                    continue

                normalized = self.clean_single_line(raw_value)

                # 例：
                # 設計品質改善 部 連絡先 計画部署 設計標準化推進 室 ...
                # 第１シート設計 部 連絡先 計画部署 室 ...
                before_label = normalized.split("計画部署", 1)[0]

                before_label = re.sub(
                    r"\s*連絡先\s*$",
                    "",
                    before_label,
                ).strip()

                matches = re.findall(
                    r"([^\s]+)\s*部(?:\s|$)",
                    before_label,
                )

                if matches:
                    candidate = self.clean_single_line(
                        matches[-1]
                    )

                    if candidate:
                        return candidate + "部"

        return ""

    def extract_department_from_text(
        self,
        text: str,
    ) -> str:
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
            match = re.search(
                pattern,
                text,
            )

            if match:
                value = self.clean_single_line(
                    match.group(1)
                )

                if (
                    value
                    and not value.endswith("部")
                ):
                    value += "部"

                return value

        return ""

    # =====================================
    # 件名
    # =====================================
    def extract_model_and_subject_from_tables(
        self,
        tables: list[list[list[str | None]]],
    ) -> tuple[str, str]:
        """
        PDF内の全テーブルから「車種」と「件名」をセットで取得する。

        pdfplumberでは帳票によって、
        ・車種と件名が別セルになる
        ・車種と件名が1セルに結合される
        ・見出しと値が同じセルの改行で保持される
        などの差異があるため、同じブロック内だけで判定する。
        """

        for table in tables:
            for row in table:
                # =====================================
                # 1セルに「車種」と「件名」が結合されている帳票
                # 例:
                #   車種 件名
                #   - 26年9月度 品質監査業務支援
                # =====================================
                for cell in row:
                    raw_value = "" if cell is None else str(cell)

                    if not ("車種" in raw_value and "件名" in raw_value):
                        continue

                    lines = [
                        self.clean_single_line(line)
                        for line in raw_value.splitlines()
                        if self.clean_single_line(line)
                    ]

                    if not lines:
                        continue

                    header_index = next(
                        (
                            index
                            for index, line in enumerate(lines)
                            if "車種" in line and "件名" in line
                        ),
                        None,
                    )

                    if (
                        header_index is not None
                        and header_index + 1 < len(lines)
                    ):
                        model_code, subject = (
                            self.split_model_and_subject_line(
                                lines[header_index + 1]
                            )
                        )

                        if subject:
                            return model_code, subject

                # =====================================
                # 「車種」と「件名」が別セルの帳票
                # =====================================
                model_index = None
                subject_index = None

                for index, cell in enumerate(row):
                    raw_value = "" if cell is None else str(cell)

                    if "車種" in raw_value and model_index is None:
                        model_index = index

                    if "件名" in raw_value and subject_index is None:
                        subject_index = index

                if model_index is None or subject_index is None:
                    continue

                model_code = self.extract_value_below_label(
                    row[model_index],
                    "車種",
                )

                subject = self.extract_value_below_label(
                    row[subject_index],
                    "件名",
                )

                if subject:
                    return (
                        self.normalize_model_code(model_code),
                        subject,
                    )

        return "", ""

    def extract_value_below_label(
        self,
        value,
        label: str,
    ) -> str:
        """
        1セル内の「見出し + 改行 + 値」から値を取得する。

        例:
            件名\n設計品質改善会の運営サポート
            → 設計品質改善会の運営サポート
        """

        if value is None:
            return ""

        lines = [
            self.clean_single_line(line)
            for line in str(value).splitlines()
            if self.clean_single_line(line)
        ]

        for index, line in enumerate(lines):
            if line == label:
                if index + 1 < len(lines):
                    return lines[index + 1].strip()

            if line.startswith(label):
                candidate = re.sub(
                    rf"^{re.escape(label)}\s*[:：]?\s*",
                    "",
                    line,
                ).strip()

                if candidate:
                    return candidate

        return ""

    def split_model_and_subject_line(
        self,
        value: str,
    ) -> tuple[str, str]:
        """
        車種と件名が同一行へ結合された値を分割する。

        例:
            - 26年9月度 品質監査業務支援
            → ("-", "26年9月度 品質監査業務支援")
        """

        cleaned = self.clean_single_line(value)

        if not cleaned:
            return "", ""

        if cleaned == "-":
            return "-", ""

        if cleaned.startswith("-"):
            return "-", cleaned[1:].strip()

        parts = cleaned.split(maxsplit=1)

        if len(parts) == 1:
            return "", parts[0]

        first, remainder = parts

        # 車種コードとして妥当な値だけを車種コード候補とする。
        # "DAS" のような英字だけの語は件名の先頭語である可能性が高いため除外する。
        if self.looks_like_model_code(first):
            return (
                self.normalize_model_code(first),
                remainder.strip(),
            )

        # 車種コードらしくなければ行全体を件名として扱う。
        return "", cleaned

    def normalize_model_code(
        self,
        value: str,
    ) -> str:
        """車種コードとして不正な見出し文字列を除外する。"""

        cleaned = self.clean_single_line(value)

        if not cleaned:
            return ""

        ignored_values = {
            "車種",
            "件名",
            "開発ﾌｪｰｽﾞ",
            "開発フェーズ",
            "部位コード",
            "シート",
        }

        if cleaned in ignored_values:
            return ""

        if (
            "開発ﾌｪｰｽﾞ" in cleaned
            or "開発フェーズ" in cleaned
        ):
            return ""

        if cleaned.startswith("件名"):
            return ""

        # 帳票で明示された「-」はそのまま有効値として扱う。
        if cleaned == "-":
            return "-"

        # "QA23" など車種コードとして妥当な値だけを残す。
        # "DAS" のような英字だけの値は誤取得として除外する。
        if not self.looks_like_model_code(cleaned):
            return ""

        return self.normalize_ascii(cleaned)

    def looks_like_model_code(
        self,
        value: str,
    ) -> bool:
        """
        車種コードらしい値か判定する。

        有効例:
            QA23
            410D
            A1X

        無効例:
            DAS
            部位コード
            シート

        現在確認しているTB帳票では車種コードは数字を1文字以上含むため、
        英字だけの語を車種コードとして扱わない。
        """

        cleaned = self.normalize_ascii(
            self.clean_single_line(value)
        )

        if not cleaned or cleaned == "-":
            return False

        if not re.fullmatch(r"[A-Za-z0-9_-]+", cleaned):
            return False

        return any(
            character.isdigit()
            for character in cleaned
        )

    def is_invalid_model_code(
        self,
        value: str,
    ) -> bool:
        """
        車種コードとして不正な値か判定する。

        空欄は不正、帳票上の「-」は有効。
        それ以外は looks_like_model_code() で判定する。
        """

        cleaned = self.clean_single_line(value)

        if not cleaned:
            return True

        if cleaned == "-":
            return False

        return not self.looks_like_model_code(cleaned)

    def is_invalid_subject(
        self,
        value: str,
    ) -> bool:
        """
        件名として明らかに不正な見出し文字列か判定する。

        PDFの表解析によって「部位コード」などの見出しが
        件名として誤取得されるケースを除外する。
        """

        cleaned = self.clean_single_line(value)

        if not cleaned:
            return True

        ignored_values = {
            "件名",
            "車種",
            "部位コード",
            "シート",
            "連絡先",
            "部",
            "室",
            "開発ﾌｪｰｽﾞ",
            "開発フェーズ",
            "OUTPUT",
            "名称",
            "部品名",
            "種類数",
            "サイズ",
            "サイズ(参考)",
        }

        if cleaned in ignored_values:
            return True

        return False


    def remove_model_code_from_subject(
        self,
        subject: str,
        model_code: str,
    ) -> str:
        """
        件名の先頭に車種コードが混入している場合に除去する。

        例:
            subject    = "QA23 無償修理低減解析"
            model_code = "QA23"

            -> "無償修理低減解析"

        車種コードと完全一致する先頭部分だけを削除するため、
        通常の件名には影響しない。
        """

        cleaned_subject = self.clean_single_line(subject)
        cleaned_model_code = self.normalize_model_code(model_code)

        if not cleaned_subject:
            return ""

        if not cleaned_model_code:
            return cleaned_subject

        pattern = (
            r"^"
            + re.escape(cleaned_model_code)
            + r"(?:\s+|　+)"
        )

        cleaned_subject = re.sub(
            pattern,
            "",
            cleaned_subject,
            count=1,
            flags=re.IGNORECASE,
        ).strip()

        return cleaned_subject

    def extract_subject_from_table(
        self,
        table: list[list[str | None]],
    ) -> str:
        """互換用。単一テーブルから件名だけを返す。"""

        _, subject = self.extract_model_and_subject_from_tables([table])
        return subject

    def extract_subject_from_text(
        self,
        text: str,
    ) -> str:
        """
        PDF全文テキストから件名を取得する。

        帳票によってPDF内部の文字順序が異なるため、
        複数パターンで取得する。

        例:
            業務委託計画書
            無償修理低減解析
            件名

        または:
            車種 件名
            - 26年9月度 品質監査業務支援
        """

        patterns = [
            (
                r"業務委託計画書\s*\n"
                r"([^\n]+?)\s*\n"
                r"件名"
            ),
            (
                r"車種\s+件名\s*\n"
                r"[^\n]+\s+(.+)"
            ),
            r"件名\s*[:：]?\s*(.+)",
        ]

        ignored_values = {
            "連絡先",
            "部位コード",
            "車種",
            "件名",
            "部",
            "室",
            "開発ﾌｪｰｽﾞ",
            "開発フェーズ",
            "シート",
        }

        for pattern in patterns:
            match = re.search(
                pattern,
                text,
            )

            if not match:
                continue

            value = self.clean_single_line(
                match.group(1)
            )

            if not value:
                continue

            if value in ignored_values:
                continue

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

                    # 同じセル内
                    same_cell = self.extract_title_number(
                        value
                    )

                    if same_cell:
                        return same_cell

                    # タイトル右隣
                    for next_cell in row[index + 1:]:
                        candidate = self.clean_cell(
                            next_cell
                        )

                        if not candidate:
                            continue

                        number = (
                            self.extract_number_candidate(
                                candidate
                            )
                        )

                        if number:
                            return number

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

            number = self.extract_title_number(
                line
            )

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

        title_position = value.find(
            "業務委託計画書"
        )

        if title_position < 0:
            return ""

        suffix = value[
            title_position
            + len("業務委託計画書"):
        ].strip()

        suffix = re.sub(
            r"^(?:No\.?|NO\.?|Ｎｏ\.?|番号)"
            r"\s*[:：]?\s*",
            "",
            suffix,
            flags=re.IGNORECASE,
        )

        return self.extract_number_candidate(
            suffix
        )

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

        if not any(
            character.isdigit()
            for character in normalized
        ):
            return ""

        return normalized

    # =====================================
    # （参考）作業内容
    # =====================================
    def extract_work_items_from_tables(
        self,
        tables: list[list[list[str | None]]],
    ) -> list[str]:
        """
        PDF内の全テーブルから作業内容を取得する。
        """

        for table in tables:
            items = self.extract_work_items_from_table(table)

            if items:
                return items

        return []

    def extract_work_items_from_table(
        self,
        table: list[list[str | None]],
    ) -> list[str]:
        """
        「（参考）作業内容」欄の項目を取得する。

        判定方針:
        ・PDF上の改行だけでは別項目と判定しない。
        ・表の行を基本単位とする。
        ・①②③、1.、(1) などの項目番号があれば新規項目とする。
        ・予定工数／予定金額が入っている行も新規項目とする。
        ・項目番号も予定値もない行は、直前項目の折り返し・続きとして結合する。
        ・見積書表示用に、各論理項目の1行目と2行目を別途保持する。

        2行目保持を追加しても、既存の新規項目／続き行判定は変更しない。
        """

        items: list[str] = []
        titles: list[str] = []
        second_lines: list[str] = []

        section_started = False
        item_column_index: int | None = None
        plan_column_index: int | None = None

        for row in table:
            cleaned_cells = [
                self.clean_cell(cell)
                for cell in row
            ]

            row_text = " ".join(
                value
                for value in cleaned_cells
                if value
            )

            # =====================================
            # 作業内容セクション開始
            # =====================================
            if not section_started:
                has_item_header = bool(
                    re.search(
                        r"項\s*目",
                        row_text,
                    )
                )

                has_plan_column = (
                    "予定工数" in row_text
                    or "予定金額" in row_text
                )

                if not (
                    has_item_header
                    and has_plan_column
                ):
                    continue

                section_started = True

                for index, cleaned in enumerate(
                    cleaned_cells
                ):
                    if re.fullmatch(
                        r"項\s*目",
                        cleaned,
                    ):
                        item_column_index = index
                        break

                for index, cleaned in enumerate(
                    cleaned_cells
                ):
                    if (
                        "予定工数" in cleaned
                        or "予定金額" in cleaned
                    ):
                        plan_column_index = index
                        break

                continue

            # =====================================
            # 作業内容セクション終了
            # =====================================
            if (
                self.row_contains(row, "委託先")
                or self.row_contains(row, "委託金額")
                or self.row_contains(row, "開始予定")
                or self.row_contains(row, "納期")
                or self.row_contains(row, "設備等の貸与")
                or self.row_contains(row, "納入場所")
                or self.row_contains(row, "支払条件")
            ):
                break

            if any(
                self.clean_single_line(cell) == "合計"
                for cell in cleaned_cells
                if cell
            ):
                break

            if (
                "作業内容" in row_text
                and len(
                    [
                        value
                        for value in cleaned_cells
                        if value
                    ]
                ) <= 2
            ):
                continue

            # =====================================
            # 項目列の生データ取得
            # =====================================
            raw_value = ""

            if (
                item_column_index is not None
                and item_column_index < len(row)
                and row[item_column_index]
            ):
                raw_value = str(
                    row[item_column_index]
                )

            if not raw_value:
                for index, candidate in enumerate(row):
                    if index == plan_column_index:
                        continue

                    candidate_text = self.clean_cell(
                        candidate
                    )

                    if not candidate_text:
                        continue

                    if self.is_ignored_work_item_value(
                        candidate_text
                    ):
                        continue

                    raw_value = str(candidate)
                    break

            if not raw_value:
                continue

            # =====================================
            # セル内の物理行を保持
            # =====================================
            raw_lines = [
                self.clean_single_line(line)
                for line in raw_value.splitlines()
                if self.clean_single_line(line)
            ]

            if not raw_lines:
                continue

            raw_joined = " ".join(raw_lines)

            if self.is_ignored_work_item_value(
                raw_joined
            ):
                continue

            has_item_marker = self.has_work_item_marker(
                raw_lines[0]
            )

            cleaned_item = self.clean_work_item(
                raw_joined
            )

            if not cleaned_item:
                continue

            first_line_title = self.clean_work_item(
                raw_lines[0]
            )

            second_line = ""

            if len(raw_lines) >= 2:
                second_line = self.clean_work_item(
                    raw_lines[1]
                )

            # =====================================
            # 予定工数／予定金額
            # =====================================
            plan_value = ""

            if (
                plan_column_index is not None
                and plan_column_index < len(row)
            ):
                plan_value = self.clean_cell(
                    row[plan_column_index]
                )

            has_plan_value = self.has_work_item_plan_value(
                plan_value
            )

            # =====================================
            # 新規項目 / 続き行
            # =====================================
            if has_item_marker or has_plan_value:
                items.append(cleaned_item)
                titles.append(
                    first_line_title
                    or cleaned_item
                )
                second_lines.append(
                    second_line
                )
                continue

            if items:
                items[-1] = self.clean_single_line(
                    f"{items[-1]} {cleaned_item}"
                )

                # 直前項目に2行目がまだ無ければ、
                # 続き行の先頭を2行目として採用する。
                if (
                    second_lines
                    and not second_lines[-1]
                ):
                    second_lines[-1] = (
                        first_line_title
                        or cleaned_item
                    )
            else:
                items.append(cleaned_item)
                titles.append(
                    first_line_title
                    or cleaned_item
                )
                second_lines.append(
                    second_line
                )

        normalized_items = (
            self.nonempty_preserve_duplicates(
                items
            )
        )
        normalized_titles = (
            self.nonempty_preserve_duplicates(
                titles
            )
        )

        if normalized_items:
            self._last_work_item_titles = (
                normalized_titles
            )

            normalized_second_lines = [
                self.clean_work_item(value)
                if value else ""
                for value in second_lines
            ]

            if (
                len(normalized_second_lines)
                < len(normalized_titles)
            ):
                normalized_second_lines.extend(
                    [
                        ""
                        for _ in range(
                            len(normalized_titles)
                            - len(normalized_second_lines)
                        )
                    ]
                )

            self._last_work_item_second_lines = (
                normalized_second_lines[
                    :len(normalized_titles)
                ]
            )

        return normalized_items

    def has_work_item_marker(
        self,
        value: str,
    ) -> bool:
        """作業項目の先頭に明示的な番号・記号があるか判定する。"""

        cleaned = self.clean_single_line(value)

        if not cleaned:
            return False

        return bool(
            re.match(
                r"^(?:"
                r"[①②③④⑤⑥⑦⑧⑨⑩]+(?:[：:])?"
                r"|\(?\d+\)?[\.．、：:]"
                r")",
                cleaned,
            )
        )

    def has_work_item_plan_value(
        self,
        value: str,
    ) -> bool:
        """予定工数／予定金額セルに実データがあるか判定する。"""

        cleaned = self.clean_single_line(value)

        if not cleaned:
            return False

        if (
            "予定工数" in cleaned
            or "予定金額" in cleaned
        ):
            return False

        # 「－」も、元帳票でその行が独立した項目であることを示す値として扱う。
        if cleaned in {"-", "－", "―", "ー"}:
            return True

        if re.fullmatch(
            r"\d+(?:\.\d+)?\s*Hr",
            cleaned,
            flags=re.IGNORECASE,
        ):
            return True

        if re.fullmatch(
            r"[￥¥]?\s*[\d,，]+(?:\.\d+)?\s*円?",
            cleaned,
        ):
            return True

        # 上記以外でも予定列に文字が入っていれば独立項目の補助情報とする。
        return True

    def is_ignored_work_item_value(
        self,
        value: str,
    ) -> bool:
        """作業内容として扱わない見出し・工数・金額を判定する。"""

        cleaned = self.clean_single_line(value)

        if not cleaned:
            return True

        if cleaned in {
            "合計",
            "項目",
            "項 目",
            "（参考）",
            "(参考)",
            "作業内容",
            "（参考）作業内容",
            "(参考)作業内容",
            "（参考）予定工数",
            "(参考)予定工数",
            "（参考）予定金額",
            "(参考)予定金額",
        }:
            return True

        if (
            "予定工数" in cleaned
            or "予定金額" in cleaned
        ):
            return True

        if re.fullmatch(
            r"\d+(?:\.\d+)?\s*Hr",
            cleaned,
            flags=re.IGNORECASE,
        ):
            return True

        if re.fullmatch(
            r"[￥¥]?\s*[\d,，]+(?:\.\d+)?\s*円?",
            cleaned,
        ):
            return True

        return False

    def extract_work_items_from_text(
        self,
        text: str,
    ) -> list[str]:
        """
        表抽出に失敗した場合の補助処理。

        次の両方へ対応する。

            （参考）作業内容

        または

            （参考）
            作業内容

        「項目」の後から、
        「合計」「委託先」「委託金額」等までを
        作業内容として取得する。
        """

        if not text:
            return []

        # =====================================
        # 作業内容セクションの開始位置
        # =====================================
        start_match = re.search(
            r"(?:"
            r"（参考）|\(参考\)"
            r")?"
            r"\s*"
            r"作業内容",
            text,
            flags=re.IGNORECASE,
        )

        if not start_match:
            return []

        section_text = text[
            start_match.end():
        ]

        # =====================================
        # 「項目」見出しより後ろを対象にする
        # =====================================
        item_header_match = re.search(
            r"項\s*目",
            section_text,
        )

        if item_header_match:
            section_text = section_text[
                item_header_match.end():
            ]

        # =====================================
        # 終了位置
        # =====================================
        end_match = re.search(
            r"\n\s*(?:"
            r"合計"
            r"|委託先"
            r"|委託金額"
            r"|開始予定"
            r"|納期"
            r"|設備等の貸与"
            r"|納入場所"
            r"|支払条件"
            r")",
            section_text,
        )

        if end_match:
            section_text = section_text[
                :end_match.start()
            ]

        items: list[str] = []

        for line in section_text.splitlines():
            cleaned = self.clean_single_line(
                line
            )

            if not cleaned:
                continue

            # =====================================
            # 行末の工数を除去
            # 例:
            # 設計改善会対応 15.0Hr
            # =====================================
            cleaned = re.sub(
                r"\s+\d+(?:\.\d+)?\s*Hr\s*$",
                "",
                cleaned,
                flags=re.IGNORECASE,
            ).strip()

            # =====================================
            # 行末の予定金額を除去
            # 例:
            # DAS解析データ整備 880,000円
            # =====================================
            cleaned = re.sub(
                r"\s+[￥¥]?"
                r"[\d,，]+(?:\.\d+)?"
                r"\s*円\s*$",
                "",
                cleaned,
            ).strip()

            cleaned = self.clean_work_item(
                cleaned
            )

            if not cleaned:
                continue

            if cleaned in (
                "合計",
                "項目",
                "項 目",
                "（参考）",
                "(参考)",
                "作業内容",
                "（参考）作業内容",
                "(参考)作業内容",
            ):
                continue

            if (
                "予定工数" in cleaned
                or "予定金額" in cleaned
            ):
                continue

            # 工数のみ
            if re.fullmatch(
                r"\d+(?:\.\d+)?\s*Hr",
                cleaned,
                flags=re.IGNORECASE,
            ):
                continue

            # 金額のみ
            if re.fullmatch(
                r"[￥¥]?\s*[\d,，]+"
                r"(?:\.\d+)?\s*円?",
                cleaned,
            ):
                continue

            items.append(cleaned)

        return self.unique_nonempty(
            items
        )

    # =====================================
    # 委託金額
    # =====================================
    def extract_amount_from_table(
        self,
        table: list[list[str | None]],
    ) -> str:
        for row in table:
            if not self.row_contains(
                row,
                "委託金額",
            ):
                continue

            row_text = " ".join(
                self.clean_cell(cell)
                for cell in row
                if self.clean_cell(cell)
            )

            amount = self.find_amount(
                row_text
            )

            if amount:
                return amount

        return ""

    def extract_amount_from_text(
        self,
        text: str,
    ) -> str:
        match = re.search(
            r"委託金額\s*"
            r"([￥¥]?\s*[\d,，]+"
            r"(?:\.\d+)?\s*円?)",
            text,
        )

        if not match:
            return ""

        return self.normalize_amount_text(
            match.group(1)
        )

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

        return self.extract_due_date_from_text(
            all_text
        )

    def extract_due_date_from_text(
        self,
        text: str,
    ) -> str:
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

        # pdfplumberの表抽出で「納期」と日付が
        # 別セルになる場合
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
        """互換用。単一テーブルから車種コードだけを返す。"""

        model_code, _ = self.extract_model_and_subject_from_tables([table])
        return model_code

    def extract_model_code_from_text(
        self,
        text: str,
    ) -> str:
        """
        PDF全文テキストから車種コードを取得する。

        「車種 件名」の次行が件名だけの場合、
        件名先頭の英字（例: DAS）を車種コードとして誤取得しない。

        対応例:
            車種 件名
            QA23 無償修理低減解析
            -> QA23

            車種 件名
            - 26年9月度 品質監査業務支援
            -> -

            車種 件名
            DASデーターのまとめ資料作成
            -> ""

            車種 件名
            設計品質改善会の運営サポート
            -
            -> -
        """

        if not text:
            return ""

        lines = [
            self.clean_single_line(line)
            for line in text.splitlines()
            if self.clean_single_line(line)
        ]

        # =====================================
        # 「車種 件名」が同じ行にある帳票
        # =====================================
        for index, line in enumerate(lines):
            if not (
                "車種" in line
                and "件名" in line
            ):
                continue

            if index + 1 >= len(lines):
                continue

            next_line = lines[index + 1]

            # "- 件名..." の形式
            if next_line == "-" or next_line.startswith("- "):
                return "-"

            # "QA23 件名..." の形式
            parts = next_line.split(maxsplit=1)

            if len(parts) >= 2:
                first = parts[0]

                if self.looks_like_model_code(first):
                    return self.normalize_model_code(first)

            # 件名だけが次行にあり、その次の行に「-」がある形式
            if (
                index + 2 < len(lines)
                and lines[index + 2] == "-"
            ):
                return "-"

            # 次行が件名だけなら車種コードは空欄
            return ""

        # =====================================
        # 「車種」の次行に値がある帳票
        # =====================================
        for index, line in enumerate(lines):
            if line != "車種":
                continue

            if index + 1 >= len(lines):
                continue

            candidate = lines[index + 1]

            if candidate == "-":
                return "-"

            if self.looks_like_model_code(candidate):
                return self.normalize_model_code(candidate)

        return ""

    # =====================================
    # 共通補助処理
    # =====================================
    def find_amount(
        self,
        text: str,
    ) -> str:
        match = re.search(
            r"([￥¥]?\s*[\d,，]+"
            r"(?:\.\d+)?\s*円)",
            text,
        )

        if not match:
            return ""

        return self.normalize_amount_text(
            match.group(1)
        )

    def normalize_amount_text(
        self,
        value: str,
    ) -> str:
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

    def clean_cell(
        self,
        value,
    ) -> str:
        if value is None:
            return ""

        text = str(
            value
        ).replace(
            "\r",
            "\n",
        )

        text = re.sub(
            r"\s+",
            " ",
            text,
        )

        return text.strip()

    def clean_single_line(
        self,
        value: str,
    ) -> str:
        if not value:
            return ""

        return re.sub(
            r"\s+",
            " ",
            str(value),
        ).strip()

    def clean_work_item(
        self,
        value: str,
    ) -> str:
        if not value:
            return ""

        cleaned = self.clean_single_line(
            value
        )

        cleaned = re.sub(
            r"^[①②③④⑤⑥⑦⑧⑨⑩]+[：:\s]*",
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
            if keyword in self.clean_cell(
                cell
            ):
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
            cleaned = self.clean_cell(
                value
            )

            if cleaned:
                return cleaned

        return ""

    def looks_like_phone_number(
        self,
        value: str,
    ) -> bool:
        return bool(
            re.fullmatch(
                r"\d{2,4}-\d{2,4}",
                value,
            )
        )

    def normalize_ascii(
        self,
        value: str,
    ) -> str:
        translation = str.maketrans(
            "ＡＢＣＤＥＦＧＨＩＪＫＬＭＮＯＰＱＲＳＴＵＶＷＸＹＺ"
            "ａｂｃｄｅｆｇｈｉｊｋｌｍｎｏｐｑｒｓｔｕｖｗｘｙｚ"
            "０１２３４５６７８９",
            "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
            "abcdefghijklmnopqrstuvwxyz"
            "0123456789",
        )

        return value.translate(
            translation
        ).strip()

    def nonempty_preserve_duplicates(
        self,
        values: list[str],
    ) -> list[str]:
        """
        空文字だけを除外し、重複は残したまま返す。

        作業項目では、同じ見出しが別項目として複数回登場することがある。
        そのため output_titles など、項目数・順序の維持が必要な用途では
        unique_nonempty() を使わず、このメソッドを使用する。
        """

        result: list[str] = []

        for value in values:
            cleaned = self.clean_single_line(value)

            if not cleaned:
                continue

            result.append(cleaned)

        return result

    def unique_nonempty(
        self,
        values: list[str],
    ) -> list[str]:
        result: list[str] = []
        seen: set[str] = set()

        for value in values:
            cleaned = self.clean_single_line(
                value
            )

            if (
                not cleaned
                or cleaned in seen
            ):
                continue

            seen.add(cleaned)
            result.append(cleaned)

        return result