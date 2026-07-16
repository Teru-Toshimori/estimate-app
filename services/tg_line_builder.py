import logging

from models.tg_ocr_result import TgOcrResult


logger = logging.getLogger(__name__)


class TgLineBuilder:
    """
    TG帳票 OCR結果を行単位へ変換

    OCR結果
        ↓
    同じ高さ(Y座標)でグループ化
        ↓
    X座標順に並べる
        ↓
    行データへ変換
    """

    # 同一行と判定するY座標の許容差(px)
    Y_THRESHOLD = 15

    def build(
        self,
        ocr_results: list[TgOcrResult]
    ) -> list[list[TgOcrResult]]:
        """
        OCR結果を行へまとめる

        Returns
        -------
        list[list[TgOcrResult]]
        """

        if not ocr_results:
            return []

        # Y → X の順でソート
        results = sorted(
            ocr_results,
            key=lambda r: (r.y, r.x)
        )

        lines: list[list[TgOcrResult]] = []

        current_line: list[TgOcrResult] = []

        current_y = results[0].y

        for item in results:

            if abs(item.y - current_y) <= self.Y_THRESHOLD:

                current_line.append(item)

            else:

                current_line.sort(
                    key=lambda r: r.x
                )

                lines.append(current_line)

                current_line = [item]

                current_y = item.y

        if current_line:

            current_line.sort(
                key=lambda r: r.x
            )

            lines.append(current_line)

        logger.info("OCR行数 : %d", len(lines))

        return lines

    def get_line_text(
        self,
        line: list[TgOcrResult]
    ) -> str:
        """
        行を文字列へ変換
        """

        return "".join(
            item.text
            for item in line
        )

    def print_lines(
        self,
        lines: list[list[TgOcrResult]]
    ) -> None:
        """
        デバッグ表示
        """

        logger.info("========== OCR行一覧 ==========")

        for i, line in enumerate(lines, start=1):

            text = self.get_line_text(line)

            logger.info(
                "%02d : %s",
                i,
                text
            )