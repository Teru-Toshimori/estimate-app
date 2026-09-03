import os
import re
import unicodedata
from datetime import datetime, timedelta
from pathlib import Path

from services.excel.excel_automation_helper import ExcelAutomationSession
from services.text.work_item_shortener import WorkItemShortener


class ExcelWriter:
    """
    見積書テンプレートへデータを書き込み、
    Excel・PDFを出力する。

    Excelはバックグラウンドで起動し、
    警告表示、リンク更新、イベント実行、
    画面更新、不要なプリンター通信を抑制する。

    対象:
        ・特調TB
        ・特調以外TB

    見積有効期限:
        発行日から14日後の日付をJ31へ記入する。
    """

    # =====================================
    # 作業内容欄
    # =====================================

    # 見積書テンプレート上で
    # 作業内容として使用できる行。
    OUTPUT_START_ROW = 18
    OUTPUT_END_ROW = 24

    # =====================================
    # 特調TB / 特調以外TB 作業内容表示設定
    # =====================================

    # 特調以外TBの作業項目タイトルを
    # 1行で表示する際の文字幅目安。
    #
    # 全角文字 = 2
    # 半角文字 = 1
    #
    # 1行目自体がこの幅を超えた場合のみ
    # 末尾を「…」で省略する。
    TOKUCHO_OTHER_LINE_WIDTH = 74

    def write(
        self,
        template_path: str,
        output_path: str,
        data,
    ) -> None:

        template_path = str(
            Path(template_path).resolve()
        )

        output_path = str(
            Path(output_path).resolve()
        )

        pdf_path = str(
            Path(output_path).with_suffix(".pdf")
        )

        self._validate_template(
            template_path
        )

        Path(output_path).parent.mkdir(
            parents=True,
            exist_ok=True,
        )

        self._remove_existing_file(
            output_path
        )

        self._remove_existing_file(
            pdf_path
        )

        session = ExcelAutomationSession()

        app = None
        template_book = None
        output_book = None
        sheet = None

        try:
            # =====================================
            # Excel起動
            # =====================================
            app = session.start()

            # =====================================
            # テンプレートを開く
            # =====================================
            template_book = app.books.open(
                template_path,
                update_links=False,
                read_only=True,
                ignore_read_only_recommended=True,
                notify=False,
                add_to_mru=False,
            )

            sheet_names = [
                worksheet.name
                for worksheet in template_book.sheets
            ]

            if "フォーマット" not in sheet_names:
                raise ValueError(
                    "テンプレートに"
                    "「フォーマット」シートが"
                    "見つかりません。\n\n"
                    f"{template_path}"
                )

            # =====================================
            # フォーマットシートを新規ブックへコピー
            # =====================================
            template_sheet = (
                template_book
                .sheets["フォーマット"]
            )

            template_sheet.copy()

            output_book = app.books.active
            sheet = output_book.sheets[0]

            # =====================================
            # シート整理
            # =====================================
            sheet.name = "見積書"

            for worksheet in list(
                output_book.sheets
            ):
                if worksheet.name != "見積書":
                    worksheet.delete()

            # =====================================
            # 発行日・見積有効期限
            # =====================================
            issue_datetime = datetime.today()

            issue_date_text = (
                issue_datetime.strftime(
                    "%Y/%m/%d"
                )
            )

            valid_until_text = (
                issue_datetime
                + timedelta(days=14)
            ).strftime(
                "%Y/%m/%d"
            )

            # =====================================
            # 基本情報
            # =====================================

            # 見積番号
            sheet.range(
                "L1"
            ).value = self._safe_text(
                data.estimate_no
            )

            # 発行日
            sheet.range(
                "K3"
            ).value = issue_date_text

            # 見積有効期限
            sheet.range(
                "J31"
            ).value = valid_until_text

            # 依頼部署
            sheet.range(
                "A7"
            ).value = self._add_onchu(
                data.department
            )

            # 件名
            sheet.range(
                "E14"
            ).value = self._safe_text(
                data.subject
            )

            # 委託金額
            sheet.range(
                "J18"
            ).value = self._remove_yen(
                data.amount
            )

            # =====================================
            # 伝票番号
            # =====================================
            voucher_no = self._safe_text(
                getattr(
                    data,
                    "voucher_no",
                    "",
                )
            )

            if voucher_no:
                sheet.range(
                    "B25"
                ).value = "伝票ＮＯ．"

                sheet.range(
                    "C25"
                ).value = voucher_no

            else:
                sheet.range(
                    "B25"
                ).value = ""

                sheet.range(
                    "C25"
                ).value = ""

            # =====================================
            # 申請書No
            # =====================================
            application_no = self._safe_text(
                getattr(
                    data,
                    "application_no",
                    "",
                )
            )

            if application_no:
                sheet.range(
                    "B26"
                ).value = "申請書ＮＯ．"

                sheet.range(
                    "C26"
                ).value = application_no

            else:
                sheet.range(
                    "B26"
                ).value = ""

                sheet.range(
                    "C26"
                ).value = ""

            # =====================================
            # 納期
            # =====================================
            sheet.range(
                "C30"
            ).value = self._safe_text(
                data.due_date
            )

            # J30はテンプレート内の文章を残し、
            # 日付部分のみ置換する。
            original = (
                sheet.range(
                    "J30"
                ).value
            )

            sheet.range(
                "J30"
            ).value = self._replace_date_in_text(
                original,
                data.due_date,
            )

            # =====================================
            # 作業内容・成果物
            # =====================================
            start_row = self.OUTPUT_START_ROW
            end_row = self.OUTPUT_END_ROW

            # =====================================
            # レイアウトモード判定
            # =====================================
            output_layout_mode = (
                self._safe_text(
                    getattr(
                        data,
                        "output_layout_mode",
                        "",
                    )
                )
                .strip()
                .lower()
            )

            # =====================================
            # 出力する作業項目を選択
            # =====================================
            #
            # 特調以外TB:
            #   Readerで保持したoutput_titles
            #   （各作業項目の1行目）を最優先する。
            #
            # 特調TB:
            #   従来どおりoutputsを使用する。
            display_output_rows = None

            if output_layout_mode in (
                "tokucho",
                "tokucho_other",
            ):
                display_output_rows = getattr(
                    data,
                    "display_output_rows",
                    None,
                )

                # 特調TBではReaderがoutput_titles / output_second_linesを
                # 保持するため、専用Writerを挟まなくてもここで
                # 特調以外TBと同じ表示行を生成できるようにする。
                if (
                    output_layout_mode == "tokucho"
                    and not display_output_rows
                ):
                    display_output_rows = (
                        self._build_extended_display_rows(
                            data=data,
                            max_rows=(
                                end_row
                                - start_row
                                + 1
                            ),
                        )
                    )

                    setattr(
                        data,
                        "display_output_rows",
                        display_output_rows,
                    )

                    setattr(
                        data,
                        "display_outputs",
                        [
                            row.get("text", "")
                            for row in display_output_rows
                            if row.get("text")
                        ],
                    )

                raw_outputs = (
                    getattr(
                        data,
                        "display_outputs",
                        None,
                    )
                    or getattr(
                        data,
                        "output_titles",
                        None,
                    )
                    or getattr(
                        data,
                        "items",
                        None,
                    )
                    or getattr(
                        data,
                        "outputs",
                        [],
                    )
                    or []
                )
            else:
                raw_outputs = (
                    getattr(
                        data,
                        "outputs",
                        [],
                    )
                    or []
                )

            outputs = [
                str(value).strip()
                for value in raw_outputs
                if (
                    value is not None
                    and str(value).strip()
                )
            ]

            # =====================================
            # 作業内容欄をクリア
            # =====================================
            self._clear_output_area(
                sheet=sheet,
                start_row=start_row,
                end_row=end_row,
            )

            if output_layout_mode in (
                "tokucho",
                "tokucho_other",
            ):
                # =====================================
                # 特調TB / 特調以外TB 共通の長文対応
                # =====================================

                # テンプレート側で
                # 「縮小して全体を表示」が設定されていると、
                # 長文セルだけ文字が極端に小さくなるため解除する。
                self._configure_tokucho_other_output_cells(
                    sheet=sheet,
                    start_row=start_row,
                    end_row=end_row,
                )

                self._write_tokucho_other_outputs(
                    sheet=sheet,
                    outputs=outputs,
                    display_rows=display_output_rows,
                    start_row=start_row,
                    end_row=end_row,
                )

            else:
                # =====================================
                # 特調TB
                # =====================================
                #
                # 既存処理をそのまま使用する。
                self._write_default_outputs(
                    sheet=sheet,
                    outputs=outputs,
                    start_row=start_row,
                    end_row=end_row,
                )

            # =====================================
            # Excel保存
            # =====================================
            output_book.save(
                output_path
            )

            if not os.path.exists(
                output_path
            ):
                raise RuntimeError(
                    "見積書Excelが"
                    "作成されませんでした。\n\n"
                    f"{output_path}"
                )

            # =====================================
            # PDF出力
            # =====================================

            # PDF出力前に
            # プリンター通信を有効へ戻す。
            session.enable_print_communication()

            sheet.api.ExportAsFixedFormat(
                Type=0,
                Filename=pdf_path,
                Quality=0,
                IncludeDocProperties=True,
                IgnorePrintAreas=False,
                OpenAfterPublish=False,
            )

            if not os.path.exists(
                pdf_path
            ):
                raise RuntimeError(
                    "見積書PDFが"
                    "作成されませんでした。\n\n"
                    f"{pdf_path}"
                )

        finally:
            # =====================================
            # 後処理
            # =====================================
            self._close_book(
                output_book
            )

            self._close_book(
                template_book
            )

            sheet = None
            output_book = None
            template_book = None
            app = None

            session.close()

    # =====================================
    # 特調TB 表示行生成
    # =====================================
    def _build_extended_display_rows(
        self,
        data,
        max_rows: int = 7,
    ) -> list[dict]:
        """
        特調TBの成果物を、特調以外TBと同じ表示ルールで
        見積書用の行データへ変換する。

        ルール:
            ・通常は各成果物の1行目を表示する。
            ・同名の1行目が複数ある場合だけ2行目で区別する。
            ・2行表示しても7行以内なら、2行目をNo.なしで追加する。
            ・7行を超える場合は、同名項目だけ1行目+2行目を結合し、
              WorkItemShortenerで1行表示用に短縮する。
            ・長い1行目もWorkItemShortenerへ渡す。
        """

        work_items = (
            getattr(
                data,
                "output_titles",
                None,
            )
            or getattr(
                data,
                "outputs",
                [],
            )
            or []
        )

        work_items = [
            self._safe_text(value).strip()
            for value in work_items
            if self._safe_text(value).strip()
        ]

        if not work_items:
            return []

        second_lines = list(
            getattr(
                data,
                "output_second_lines",
                [],
            )
            or []
        )

        if len(second_lines) < len(work_items):
            second_lines.extend(
                [
                    ""
                    for _ in range(
                        len(work_items)
                        - len(second_lines)
                    )
                ]
            )

        second_lines = [
            self._safe_text(value).strip()
            for value in second_lines[:len(work_items)]
        ]

        normalized_titles = [
            self._normalize_output_title_key(value)
            for value in work_items
        ]

        title_counts: dict[str, int] = {}

        for key in normalized_titles:
            if key:
                title_counts[key] = (
                    title_counts.get(key, 0)
                    + 1
                )

        duplicate_flags = [
            bool(
                key
                and title_counts.get(key, 0) > 1
            )
            for key in normalized_titles
        ]

        extra_rows = sum(
            1
            for is_duplicate, second_line in zip(
                duplicate_flags,
                second_lines,
            )
            if is_duplicate and second_line
        )

        can_use_continuation_rows = (
            len(work_items) + extra_rows
            <= max_rows
        )

        shortener = WorkItemShortener()
        rows: list[dict] = []

        for index, title in enumerate(work_items):
            second_line = second_lines[index]
            is_duplicate = duplicate_flags[index]

            if (
                is_duplicate
                and second_line
                and can_use_continuation_rows
            ):
                rows.append(
                    {
                        "number": index + 1,
                        "text": shortener.shorten(title),
                    }
                )

                rows.append(
                    {
                        "number": None,
                        "text": shortener.shorten(
                            second_line
                        ),
                    }
                )

                continue

            display_text = title

            if is_duplicate and second_line:
                display_text = (
                    f"{title} {second_line}"
                ).strip()

            rows.append(
                {
                    "number": index + 1,
                    "text": shortener.shorten(
                        display_text
                    ),
                }
            )

        return rows[:max_rows]

    def _normalize_output_title_key(
        self,
        value,
    ) -> str:
        """空白差だけの同名成果物も同じタイトルとして扱う。"""

        text = self._safe_text(value).strip()
        return re.sub(
            r"\s+",
            " ",
            text,
        )

    # =====================================
    # 作業内容欄クリア
    # =====================================
    def _clear_output_area(
        self,
        sheet,
        start_row: int,
        end_row: int,
    ) -> None:

        for row in range(
            start_row,
            end_row + 1,
        ):
            sheet.range(
                f"A{row}"
            ).value = None

            sheet.range(
                f"B{row}"
            ).value = None

    # =====================================
    # 特調TB / 特調以外TB 作業内容セル設定
    # =====================================
    def _configure_tokucho_other_output_cells(
        self,
        sheet,
        start_row: int,
        end_row: int,
    ) -> None:
        """
        特調以外TBの作業内容セルについて、
        テンプレート側の自動縮小等を解除する。

        今回はPython側で表示文字数を制御するため、
        Excel側では文字を勝手に縮小させない。
        """

        for row in range(
            start_row,
            end_row + 1,
        ):
            cell = sheet.range(
                f"B{row}"
            )

            try:
                # 「縮小して全体を表示」をOFF
                cell.api.ShrinkToFit = False
            except Exception:
                pass

            try:
                # Python側で改行を行単位に制御するため、
                # Excelの自動折り返しは使用しない。
                cell.api.WrapText = False
            except Exception:
                pass

    # =====================================
    # 従来の成果物出力
    # =====================================
    def _write_default_outputs(
        self,
        sheet,
        outputs: list[str],
        start_row: int,
        end_row: int,
    ) -> None:
        """
        特調TBなどで使用する従来処理。

        1項目 = 1Excel行。

        特調TBの既存動作を変更しない。
        """

        for index, output in enumerate(
            outputs
        ):
            row = (
                start_row
                + index
            )

            if row > end_row:
                break

            sheet.range(
                f"A{row}"
            ).value = (
                index + 1
            )

            sheet.range(
                f"B{row}"
            ).value = self._remove_number(
                output
            )

    # =====================================
    # 特調TB / 特調以外TB 作業内容出力
    # =====================================
    def _write_tokucho_other_outputs(
        self,
        sheet,
        outputs: list[str],
        display_rows,
        start_row: int,
        end_row: int,
    ) -> None:
        """
        特調以外TB専用の作業内容表示。

        display_rows がある場合は、
        number=None の行を直前項目の続きとして
        No.欄を空白にして出力する。

        display_rows がない既存データでは、
        従来どおり outputs を1項目1行で出力する。
        """

        total_rows = (
            end_row
            - start_row
            + 1
        )

        rows_to_write: list[dict] = []

        if isinstance(
            display_rows,
            list,
        ):
            for row_data in display_rows:
                if not isinstance(
                    row_data,
                    dict,
                ):
                    continue

                text = self._remove_number(
                    row_data.get(
                        "text",
                        "",
                    )
                )

                if not text:
                    continue

                rows_to_write.append(
                    {
                        "number": (
                            row_data.get(
                                "number",
                                None,
                            )
                        ),
                        "text": text,
                    }
                )

        # 互換フォールバック
        if not rows_to_write:
            cleaned_outputs = []

            for output in outputs:
                cleaned = self._remove_number(
                    output
                )

                if cleaned:
                    cleaned_outputs.append(
                        cleaned
                    )

            rows_to_write = [
                {
                    "number": index + 1,
                    "text": text,
                }
                for index, text in enumerate(
                    cleaned_outputs
                )
            ]

        rows_to_write = rows_to_write[
            :total_rows
        ]

        for offset, row_data in enumerate(
            rows_to_write
        ):
            row = (
                start_row
                + offset
            )

            item_number = row_data.get(
                "number",
                None,
            )

            sheet.range(
                f"A{row}"
            ).value = (
                item_number
                if item_number is not None
                else None
            )

            target_cell = sheet.range(
                f"B{row}"
            )

            try:
                target_cell.api.ShrinkToFit = False
            except Exception:
                pass

            try:
                target_cell.api.WrapText = False
            except Exception:
                pass

            target_cell.value = row_data.get(
                "text",
                "",
            )

    # =====================================
    # 8項目以上の場合
    # =====================================
    def _combine_remaining_items(
        self,
        values: list[str],
    ) -> str:
        """
        7行を超える項目がある場合、
        残りを最終行へまとめる。

        最終的な文字数制御は
        _split_text_to_lines()で行う。
        """

        cleaned = []

        for value in values:
            item = self._remove_number(
                value
            )

            if item:
                cleaned.append(
                    item
                )

        return " ／ ".join(
            cleaned
        )

    # =====================================
    # 必要行数推定
    # =====================================
    def _estimate_required_lines(
        self,
        text: str,
        line_width: int,
    ) -> int:
        """
        表示幅から必要行数を概算する。

        全角:
            2

        半角:
            1
        """

        if not text:
            return 0

        width = (
            self._text_display_width(
                text
            )
        )

        if width <= 0:
            return 1

        return max(
            1,
            (
                width
                + line_width
                - 1
            )
            // line_width,
        )

    # =====================================
    # テキスト分割
    # =====================================
    def _split_text_to_lines(
        self,
        text: str,
        line_width: int,
        max_lines: int,
    ) -> list[str]:
        """
        指定された表示幅と行数に合わせて
        テキストを分割する。

        全文が収まらない場合のみ、
        最終行末尾へ「…」を付ける。
        """

        text = (
            self._safe_text(
                text
            ).strip()
        )

        if (
            not text
            or max_lines <= 0
        ):
            return []

        remaining = text
        result: list[str] = []

        for line_index in range(
            max_lines
        ):

            if not remaining:
                break

            is_last_line = (
                line_index
                == max_lines - 1
            )

            # =====================================
            # この行で全文が収まる
            # =====================================
            if (
                self._text_display_width(
                    remaining
                )
                <= line_width
            ):
                result.append(
                    remaining.strip()
                )

                remaining = ""
                break

            # =====================================
            # 最終行
            # =====================================
            if is_last_line:

                ellipsis = "…"

                ellipsis_width = (
                    self._text_display_width(
                        ellipsis
                    )
                )

                available_width = max(
                    1,
                    line_width
                    - ellipsis_width,
                )

                chunk, _ = (
                    self._take_text_by_width(
                        remaining,
                        available_width,
                    )
                )

                chunk = chunk.rstrip(
                    " 、,，。"
                )

                result.append(
                    chunk
                    + ellipsis
                )

                remaining = ""
                break

            # =====================================
            # 次行へ続く
            # =====================================
            chunk, remaining = (
                self._take_text_by_width(
                    remaining,
                    line_width,
                )
            )

            chunk = chunk.strip()

            if chunk:
                result.append(
                    chunk
                )

            remaining = (
                remaining.lstrip()
            )

        return result

    # =====================================
    # 指定幅で文章を切り出す
    # =====================================
    def _take_text_by_width(
        self,
        text: str,
        max_width: int,
    ) -> tuple[str, str]:
        """
        max_width以内で文章を切り出す。

        可能であれば、
        空白・句読点付近で改行する。
        """

        if not text:
            return "", ""

        current_width = 0
        split_index = 0

        preferred_split_index = None

        for index, character in enumerate(
            text
        ):

            char_width = (
                self._character_display_width(
                    character
                )
            )

            if (
                current_width
                + char_width
                > max_width
            ):
                break

            current_width += char_width
            split_index = index + 1

            # 自然な改行候補
            if character in (
                " ",
                "　",
                "、",
                "。",
                "，",
                ",",
                "）",
                ")",
                "＞",
                ">",
                "：",
                ":",
            ):
                preferred_split_index = (
                    index + 1
                )

        if split_index <= 0:
            split_index = 1

        # =====================================
        # 句読点が近くにある場合は
        # そこで切る
        # =====================================
        if (
            preferred_split_index
            is not None
            and preferred_split_index
            >= max(
                1,
                int(
                    split_index
                    * 0.65
                ),
            )
        ):
            split_index = (
                preferred_split_index
            )

        return (
            text[:split_index],
            text[split_index:],
        )

    # =====================================
    # 文字列の表示幅
    # =====================================
    def _text_display_width(
        self,
        text: str,
    ) -> int:

        return sum(
            self._character_display_width(
                character
            )
            for character in str(
                text
            )
        )

    # =====================================
    # 1文字の表示幅
    # =====================================
    def _character_display_width(
        self,
        character: str,
    ) -> int:
        """
        日本語・全角文字:
            2

        半角英数字:
            1
        """

        if character == "\t":
            return 4

        east_asian_width = (
            unicodedata.east_asian_width(
                character
            )
        )

        if east_asian_width in (
            "W",
            "F",
            "A",
        ):
            return 2

        return 1

    # =====================================
    # テンプレート確認
    # =====================================
    def _validate_template(
        self,
        template_path: str,
    ) -> None:

        if not os.path.exists(
            template_path
        ):
            raise FileNotFoundError(
                "見積書テンプレートが"
                "見つかりません。\n\n"
                f"{template_path}"
            )

        if not os.path.isfile(
            template_path
        ):
            raise ValueError(
                "見積書テンプレートに"
                "ファイルを指定してください。\n\n"
                f"{template_path}"
            )

    # =====================================
    # 既存ファイル削除
    # =====================================
    def _remove_existing_file(
        self,
        file_path: str,
    ) -> None:

        if not os.path.exists(
            file_path
        ):
            return

        try:
            os.remove(
                file_path
            )

        except PermissionError as error:
            raise PermissionError(
                "既存の出力ファイルを"
                "削除できません。\n"
                "ExcelまたはPDFビューアーで"
                "開かれていないか"
                "確認してください。\n\n"
                f"{file_path}"
            ) from error

    # =====================================
    # ブックを安全に閉じる
    # =====================================
    def _close_book(
        self,
        book,
    ) -> None:

        if book is None:
            return

        try:
            book.close()

        except Exception:
            pass

    # =====================================
    # 御中追加
    # =====================================
    def _add_onchu(
        self,
        department,
    ) -> str:

        department_text = (
            self._safe_text(
                department
            ).strip()
        )

        if not department_text:
            return ""

        if department_text.endswith(
            "御中"
        ):
            return department_text

        return (
            department_text
            + "　御中"
        )

    # =====================================
    # 円記号削除
    # =====================================
    def _remove_yen(
        self,
        amount,
    ) -> str:

        amount_text = (
            self._safe_text(
                amount
            )
        )

        return (
            amount_text
            .replace(
                "円",
                "",
            )
            .replace(
                "￥",
                "",
            )
            .replace(
                "¥",
                "",
            )
            .strip()
        )

    # =====================================
    # 成果物番号削除
    # =====================================
    def _remove_number(
        self,
        text,
    ) -> str:

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
    def _replace_date_in_text(
        self,
        text,
        due_date,
    ) -> str:

        due_date_text = (
            self._safe_text(
                due_date
            )
        )

        if text is None:
            return due_date_text

        text = str(
            text
        )

        patterns = (
            r"\d{4}年\d{1,2}月\d{1,2}日",
            r"\d{4}/\d{1,2}/\d{1,2}",
            r"\d{4}-\d{1,2}-\d{1,2}",
        )

        for pattern in patterns:

            if re.search(
                pattern,
                text,
            ):
                return re.sub(
                    pattern,
                    due_date_text,
                    text,
                    count=1,
                )

        return text

    # =====================================
    # 安全な文字列変換
    # =====================================
    def _safe_text(
        self,
        value,
    ) -> str:

        if value is None:
            return ""

        return str(
            value
        )