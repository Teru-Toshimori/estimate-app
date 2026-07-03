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

        data.application_no = self.find_next(lines, "申請書NO.")
        data.department = self.find_next(lines, "依頼部署")
        data.subject = self.find_next(lines, "件名")

        data.amount = self.get_amount(lines)
        data.due_date = self.get_due_date(lines)

        data.outputs = self.get_outputs(lines)

        return data

    def find_next(self, lines, keyword):

        for i, line in enumerate(lines):

            if line.strip() == keyword:

                if i + 1 < len(lines):

                    return lines[i + 1].strip()

        return ""

    def get_amount(self, lines):

        for line in lines:

            match = re.search(r"委託金額\s+([\d,]+円)", line)

            if match:
                return match.group(1)

        return ""

    def get_due_date(self, lines):

        for line in lines:

            match = re.search(r"納期\s+(\d{4}年\d{2}月\d{2}日)", line)

            if match:
                return match.group(1)

        return ""

    def get_outputs(self, lines):

        outputs = []

        start = False

        for line in lines:

            line = line.strip()

            if line.startswith("成果物名称"):
                start = True
                continue

            if not start:
                continue

            if line.startswith("成果物に"):
                break

            if line.startswith("備考"):
                continue

            if len(line) >= 2 and line[0] in "①②③④⑤⑥":

                if len(line) <= 2:
                    continue

                outputs.append(line)

        return outputs