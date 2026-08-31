import json
import os
import re
import unicodedata
from pathlib import Path
from typing import Any


class WorkItemShortener:
    """
    特調以外TBの見積書に表示する作業項目名を短縮する。

    優先順位:
        1. そのまま表示可能なら変更しない
        2. キャッシュに過去の短縮結果があれば使用
        3. 短縮辞書に一致すれば固定の短縮名を使用
        4. OpenAI APIで見積書用の短い作業名を生成
        5. AIが利用できない・失敗した場合は元の文字列を返す

    元の作業内容全文やoutput_titles自体は変更しない。
    """

    DEFAULT_MAX_DISPLAY_WIDTH = 68
    DEFAULT_MODEL = "gpt-5-mini"

    def __init__(
        self,
        dictionary_path: str | Path | None = None,
        max_display_width: int = DEFAULT_MAX_DISPLAY_WIDTH,
    ):
        self.max_display_width = max_display_width
        self.dictionary_path = (
            Path(dictionary_path)
            if dictionary_path
            else self._resolve_dictionary_path()
        )
        self.cache_path = self._resolve_cache_path()
        self.short_name_rules = self._load_json_dict(self.dictionary_path)
        self.cache = self._load_json_dict(self.cache_path)

    def shorten_many(self, values: list[str]) -> list[str]:
        result: list[str] = []

        for value in values:
            cleaned = self._normalize_text(value)

            if not cleaned:
                continue

            result.append(self.shorten(cleaned))

        return result

    def shorten(self, value: str) -> str:
        original = self._normalize_text(value)

        if not original:
            return ""

        # 1. そのまま表示可能
        if self._fits(original):
            return original

        # 2. キャッシュ
        cached = self._normalize_text(
            self.cache.get(original, "")
        )

        if cached and self._fits(cached):
            return cached

        # 3. 固定辞書
        dictionary_value = self._find_dictionary_value(original)

        if dictionary_value:
            self._save_cache_value(
                original,
                dictionary_value,
            )
            return dictionary_value

        # 4. AI短縮
        ai_value = self._shorten_with_ai(original)

        if ai_value:
            self._save_cache_value(
                original,
                ai_value,
            )
            return ai_value

        # 5. AI失敗時は元文を保持
        # 「…」で途中切りしない。
        return original

    def _find_dictionary_value(self, original: str) -> str:
        exact = self._normalize_text(
            self.short_name_rules.get(original, "")
        )

        if exact and self._fits(exact):
            return exact

        keys = sorted(
            self.short_name_rules.keys(),
            key=len,
            reverse=True,
        )

        for key in keys:
            normalized_key = self._normalize_text(key)

            if not normalized_key:
                continue

            if original.startswith(normalized_key):
                value = self._normalize_text(
                    self.short_name_rules.get(key, "")
                )

                if value and self._fits(value):
                    return value

        return ""

    def _shorten_with_ai(self, original: str) -> str:
        if not os.getenv("OPENAI_API_KEY", "").strip():
            return ""

        try:
            from openai import OpenAI
        except Exception:
            return ""

        model = (
            os.getenv("OPENAI_WORK_ITEM_MODEL", "").strip()
            or os.getenv("OPENAI_MODEL", "").strip()
            or self.DEFAULT_MODEL
        )

        try:
            client = OpenAI()

            candidate = self._request_ai_short_name(
                client=client,
                model=model,
                original=original,
                target_chars=28,
            )

            if candidate and self._fits(candidate):
                return candidate

            # まだ長い場合のみ、さらに短く再生成
            candidate = self._request_ai_short_name(
                client=client,
                model=model,
                original=original,
                target_chars=22,
            )

            if candidate and self._fits(candidate):
                return candidate

        except Exception:
            return ""

        return ""

    def _request_ai_short_name(
        self,
        client: Any,
        model: str,
        original: str,
        target_chars: int,
    ) -> str:
        prompt = f"""
あなたは業務委託の見積書に記載する「作業項目名」を整える担当です。

以下の元の作業内容を、見積書の一覧で意味が伝わる短い作業項目名にしてください。

制約:
- 日本語で1行だけ返す
- {target_chars}文字程度を目安にする
- 元の意味を変えない
- 元文にない情報を追加しない
- 固有名詞、システム名、車種名など重要な語は可能な限り残す
- 「…」や「...」で文章を途中切りしない
- 説明文ではなく名詞句・作業名としてまとめる
- 「作業項目：」などのラベル、引用符、箇条書き記号は付けない

元の作業内容:
{original}
""".strip()

        response = client.responses.create(
            model=model,
            input=prompt,
            store=False,
        )

        return self._normalize_ai_result(
            getattr(response, "output_text", "")
        )

    def _normalize_ai_result(self, value: str) -> str:
        text = self._normalize_text(value)

        if not text:
            return ""

        first_line = text.splitlines()[0].strip()

        first_line = re.sub(
            r"^(?:作業項目|作業名|短縮名)\s*[:：]\s*",
            "",
            first_line,
        ).strip()

        first_line = first_line.strip(
            "\"'「」『』"
        )

        # 途中切りされた結果は採用しない。
        if (
            first_line.endswith("…")
            or first_line.endswith("...")
        ):
            return ""

        return first_line

    def _fits(self, value: str) -> bool:
        return (
            self._text_display_width(value)
            <= self.max_display_width
        )

    def _text_display_width(self, value: str) -> int:
        return sum(
            self._character_display_width(char)
            for char in value
        )

    def _character_display_width(self, char: str) -> int:
        if char == "\t":
            return 4

        width = unicodedata.east_asian_width(char)

        if width in ("W", "F", "A"):
            return 2

        return 1

    def _load_json_dict(self, path: Path) -> dict[str, str]:
        if not path.exists():
            return {}

        try:
            with path.open("r", encoding="utf-8") as file:
                data = json.load(file)

            if not isinstance(data, dict):
                return {}

            return {
                str(key): str(value)
                for key, value in data.items()
                if str(key).strip() and str(value).strip()
            }

        except Exception:
            return {}

    def _save_cache_value(
        self,
        original: str,
        shortened: str,
    ) -> None:
        original = self._normalize_text(original)
        shortened = self._normalize_text(shortened)

        if not original or not shortened:
            return

        self.cache[original] = shortened

        try:
            self.cache_path.parent.mkdir(
                parents=True,
                exist_ok=True,
            )

            with self.cache_path.open(
                "w",
                encoding="utf-8",
            ) as file:
                json.dump(
                    self.cache,
                    file,
                    ensure_ascii=False,
                    indent=2,
                )

        except Exception:
            pass

    def _resolve_dictionary_path(self) -> Path:
        try:
            import sys

            base_path = Path(
                getattr(
                    sys,
                    "_MEIPASS",
                    Path(__file__).resolve().parents[2],
                )
            )
        except Exception:
            base_path = Path(
                __file__
            ).resolve().parents[2]

        candidate = (
            base_path
            / "resources"
            / "work_item_short_names.json"
        )

        if candidate.exists():
            return candidate

        return (
            Path(__file__)
            .resolve()
            .parents[2]
            / "resources"
            / "work_item_short_names.json"
        )

    def _resolve_cache_path(self) -> Path:
        local_app_data = os.getenv(
            "LOCALAPPDATA",
            ""
        ).strip()

        if local_app_data:
            base = Path(local_app_data)
        else:
            base = Path.home() / ".estimate_tool"

        return (
            base
            / "EstimateTool"
            / "work_item_shortener_cache.json"
        )

    def _normalize_text(self, value: Any) -> str:
        if value is None:
            return ""

        text = str(value).replace("\r", "\n")
        text = re.sub(r"[\n\t]+", " ", text)
        text = re.sub(r"\s+", " ", text)

        return text.strip()
