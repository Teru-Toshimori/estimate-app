import re
import pdfplumber

from models.estimate_data import EstimateData


class PDFReader:

    def read_text(self, pdf_path):

        text = ""

        with pdfplumber.open(pdf_path) as pdf:

            for page in pdf.pages:

                page_text = page.extract_text()

                if page_text:
                    text += page_text + "\n"

        return text

    def parse(self, pdf_path):

        text = self.read_text(pdf_path)

        lines = text.splitlines()

        data = EstimateData()

        # 基本情報
        data.application_no = self.find_next(lines, "申請書NO.")
        data.department = self.find_next(lines, "依頼部署")
        data.subject = self.find_next(lines, "件名")

        # 車種コード
        data.model_code = self.get_model_code(lines)

        # 金額・納期
        data.amount = self.get_amount(lines)
        data.due_date = self.get_due_date(lines)

        # 成果物
        data.outputs = self.get_outputs(lines)

        return data

    # -----------------------------
    # 指定文字列の次の行を取得
    # -----------------------------
    def find_next(self, lines, keyword):

        for i, line in enumerate(lines):

            if line.strip() == keyword:

                if i + 1 < len(lines):
                    return lines[i + 1].strip()

        return ""

    # -----------------------------
    # 車種コード取得
    # -----------------------------
    def get_model_code(self, lines):

        for i, line in enumerate(lines):

            if line.strip().startswith("車種コード"):

                if i + 1 < len(lines):

                    next_line = lines[i + 1].strip()

                    # 例
                    # 410D 正式図
                    # ↓
                    # 410D
                    return next_line.split()[0]

        return ""

    # -----------------------------
    # 委託金額取得
    # -----------------------------
    def get_amount(self, lines):

        for line in lines:

            m = re.search(
                r"委託金額\s+([\d,]+円)",
                line
            )

            if m:
                return m.group(1)

        return ""

    # -----------------------------
    # 納期取得
    # -----------------------------
    def get_due_date(self, lines):

        for line in lines:

            m = re.search(
                r"納期\s+(\d{4}年\d{2}月\d{2}日)",
                line
            )

            if m:
                return m.group(1)

            m = re.search(
                r"納期\s+(\d{4}/\d{2}/\d{2})",
                line
            )

            if m:
                return m.group(1)

        return ""

    # -----------------------------
    # 成果物取得
    # -----------------------------
    def get_outputs(self, lines):

        outputs = []

        start = False

        for line in lines:

            if "成果物名称" in line:
                start = True
                continue

            if not start:
                continue

            # 成果物終了
            if (
                "成果物に" in line
                or "Requirement" in line
                or "TBから提供する" in line
            ):
                break

            # ①～⑳で始まる行のみ取得
            if re.match(r"^[①-⑳]", line.strip()):

                text = line.strip()

                # 空欄（①だけ等）は除外
                if len(text) > 1:
                    outputs.append(text)

        return outputs