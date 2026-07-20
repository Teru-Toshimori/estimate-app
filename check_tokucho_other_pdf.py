from __future__ import annotations

import sys
from pathlib import Path
from tkinter import Tk, filedialog

from services.tokucho_other_pdf_reader import TokuchoOtherPdfReader


def select_pdf() -> Path | None:
    """確認対象のPDFをダイアログから選択する。"""
    root = Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    selected = filedialog.askopenfilename(
        title="確認する特調TB以外のPDFを選択",
        filetypes=[("PDFファイル", "*.pdf")],
    )

    root.destroy()

    if not selected:
        return None

    return Path(selected)


def print_result(pdf_path: Path) -> None:
    """PDFから抽出した内容を見やすく表示する。"""
    reader = TokuchoOtherPdfReader()
    data = reader.parse(str(pdf_path))

    print("=" * 70)
    print("特調TB以外 PDF抽出結果")
    print("=" * 70)
    print(f"対象PDF             : {pdf_path.name}")
    print(f"計画部署            : {data.department}")
    print(f"件名                : {data.subject}")
    print(f"業務委託計画書No    : {data.application_no}")
    print(f"委託金額            : {data.amount}")
    print(f"納期                : {data.due_date}")
    print(f"車種コード          : {data.model_code}")
    print("-" * 70)
    print("作業内容")

    outputs = getattr(data, "outputs", []) or []

    if outputs:
        for index, item in enumerate(outputs, start=1):
            print(f"  {index}. {item}")
    else:
        print("  （取得できませんでした）")

    print("=" * 70)


def main() -> int:
    try:
        pdf_path = select_pdf()

        if pdf_path is None:
            print("PDFの選択がキャンセルされました。")
            return 0

        if not pdf_path.exists():
            print(f"PDFファイルが見つかりません: {pdf_path}")
            return 1

        print_result(pdf_path)
        return 0

    except Exception as error:
        print("\n抽出確認中にエラーが発生しました。")
        print("-" * 70)
        print(str(error))
        print("-" * 70)
        return 1


if __name__ == "__main__":
    sys.exit(main())
