import base64
import json
import logging
import os
from pathlib import Path

import fitz
from dotenv import load_dotenv
from openai import OpenAI


load_dotenv()

logger = logging.getLogger(__name__)


OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")

if not OPENAI_API_KEY:
    raise RuntimeError(
        "OPENAI_API_KEY が設定されていません。\n"
        ".env ファイルを確認してください。"
    )


client = OpenAI(
    api_key=OPENAI_API_KEY
)


def pdf_pages_to_base64_images(
    pdf_path: str | Path,
    max_pages: int = 2,
    zoom: float = 2.0,
) -> list[str]:
    """
    PDFの各ページをPNG画像に変換し、
    Base64文字列のリストとして返す。

    Parameters
    ----------
    pdf_path:
        読み込むPDFのパス。

    max_pages:
        OpenAIへ送信する最大ページ数。
        業務委託計画書は通常2ページのため、初期値は2。

    zoom:
        PDFを画像化する際の拡大率。
        値を大きくすると読み取り精度が上がるが、
        データ容量も大きくなる。
    """

    path = Path(pdf_path)

    if not path.exists():
        raise FileNotFoundError(
            f"PDFファイルが見つかりません。\n{path}"
        )

    image_base64_list: list[str] = []

    document = None

    try:
        document = fitz.open(str(path))

        page_count = min(
            len(document),
            max_pages,
        )

        matrix = fitz.Matrix(
            zoom,
            zoom,
        )

        for page_index in range(page_count):
            page = document.load_page(
                page_index
            )

            pixmap = page.get_pixmap(
                matrix=matrix,
                alpha=False,
            )

            png_bytes = pixmap.tobytes(
                "png"
            )

            image_base64 = base64.b64encode(
                png_bytes
            ).decode(
                "utf-8"
            )

            image_base64_list.append(
                image_base64
            )

    except Exception as error:
        logger.exception(
            "PDF画像変換失敗: %s",
            path,
        )

        raise RuntimeError(
            "PDFを画像へ変換できませんでした。\n\n"
            f"対象PDF：{path.name}\n"
            f"詳細：{error}"
        ) from error

    finally:
        if document is not None:
            document.close()

    if not image_base64_list:
        raise ValueError(
            "PDFからページ画像を取得できませんでした。\n\n"
            f"対象PDF：{path.name}"
        )

    return image_base64_list


def parse_json_response(
    content: str,
) -> dict:
    """
    OpenAIから返されたJSON文字列をdictへ変換する。
    """

    try:
        data = json.loads(content)

    except json.JSONDecodeError as error:
        logger.exception(
            "OpenAI返却JSONの解析失敗: %s",
            content,
        )

        raise ValueError(
            "AIの解析結果をJSONとして読み取れませんでした。"
        ) from error

    outputs = data.get(
        "outputs",
        [],
    )

    if not isinstance(outputs, list):
        outputs = []

    cleaned_outputs: list[str] = []

    for output in outputs:
        value = str(output).strip()

        if value:
            cleaned_outputs.append(
                value
            )

    return {
        "application_no": str(
            data.get(
                "application_no",
                "",
            )
        ).strip(),
        "department": str(
            data.get(
                "department",
                "",
            )
        ).strip(),
        "subject": str(
            data.get(
                "subject",
                "",
            )
        ).strip(),
        "model_code": str(
            data.get(
                "model_code",
                "",
            )
        ).strip(),
        "amount": str(
            data.get(
                "amount",
                "",
            )
        ).strip(),
        "due_date": str(
            data.get(
                "due_date",
                "",
            )
        ).strip(),
        "outputs": cleaned_outputs,
    }


def normalize_department(
    department: str,
) -> str:
    """
    依頼部署名に混入した半角・全角スペースを除去する。

    例
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


def extract_tokucho_pdf_data(
    pdf_path: str | Path,
) -> dict:
    """
    画像PDFの業務委託計画書をOpenAI Visionで解析する。

    Returns
    -------
    {
        "application_no": str,
        "department": str,
        "subject": str,
        "model_code": str,
        "amount": str,
        "due_date": str,
        "outputs": list[str],
    }
    """

    path = Path(pdf_path)

    logger.info(
        "特調TB画像PDFのAI解析開始: %s",
        path.name,
    )

    image_base64_list = pdf_pages_to_base64_images(
        pdf_path=path,
        max_pages=2,
        zoom=2.0,
    )

    prompt = """
次の画像は、日本語と英語が併記された
「業務委託計画書（Outsourcing Planning Sheet）」です。

画像を読み取り、以下の項目だけを正確に抽出してください。

【抽出項目】

1. application_no
「申請書NO.」または「Application No.」の値。
例：ITK20309

2. department
「依頼部署」または「Request Div.」の値。
英語見出しではなく、実際の部署名を取得してください。
例：Rrシート骨格設計部

部署名のアルファベット間に、
帳票レイアウトによる不要な空白がある場合は除去してください。

例：
R rシート骨格設計部
→ Rrシート骨格設計部

F rシート骨格設計部
→ Frシート骨格設計部

3. subject
「件名」または「Job Title」の値。

4. model_code
「車種コード」または「Model Code」の値。

5. amount
2ページ目の「委託金額」または
「Outsourcing Job Fee」の値。
通貨単位の「円」まで含めてください。
例：5,292,000円

6. due_date
2ページ目の「納期」または
「Delivery Due date」の値。
次の形式で返してください。
YYYY年MM月DD日

例：2026年08月24日

7. outputs
1ページ目の表にある「成果物名称」または
「Name of output(s)」列に記載された成果物名称だけの一覧。

重要：
・成果物表では「成果物名称」と「備考（数量／サイズ指定など）」
  または「Remarks (Volume/Size etc.)」は別の列です。
・outputsには「成果物名称」列の文字だけを入れてください。
・右側の「備考」列の内容はoutputsへ絶対に含めないでください。
・備考が「一式」「一式／一任」「1/A3」「10枚／A4」
  「数量、サイズ指定無」などであっても、すべて除外してください。
・成果物名称が複数行で記載されている場合は、
  同じ成果物セル内の文字を1つの成果物として取得してください。
・成果物番号の①、②、③などは付けないでください。
・空欄の成果物行は含めないでください。

例1：
成果物名称：構造計画図
備考：10枚／A4
→ outputs = ["構造計画図"]

例2：
成果物名称：設変図面：内装ALL
備考：一式／一任
→ outputs = ["設変図面：内装ALL"]

例3：
成果物名称：2026年度VE活動 第1L検討書
備考：1/A3
→ outputs = ["2026年度VE活動 第1L検討書"]

例4：
成果物名称：Rrシート骨格構造計画図、検討図、3D
備考：数量、サイズ指定無
→ outputs = ["Rrシート骨格構造計画図、検討図、3D"]

【注意事項】

・帳票の見出し文字を値として取得しないでください。
・存在しない項目は空文字または空配列にしてください。
・推測で値を作らないでください。
・必ず次のJSON形式だけで返してください。

{
  "application_no": "",
  "department": "",
  "subject": "",
  "model_code": "",
  "amount": "",
  "due_date": "",
  "outputs": []
}
"""

    content_parts: list[dict] = [
        {
            "type": "text",
            "text": prompt,
        }
    ]

    for image_base64 in image_base64_list:
        content_parts.append(
            {
                "type": "image_url",
                "image_url": {
                    "url": (
                        "data:image/png;base64,"
                        f"{image_base64}"
                    ),
                    "detail": "high",
                },
            }
        )

    try:
        response = client.chat.completions.create(
            model="gpt-4.1-mini",
            response_format={
                "type": "json_object",
            },
            messages=[
                {
                    "role": "user",
                    "content": content_parts,
                }
            ],
            temperature=0,
        )

        content = (
            response
            .choices[0]
            .message
            .content
        )

        if not content:
            raise ValueError(
                "AIから解析結果が返されませんでした。"
            )

        logger.info(
            "特調TB画像PDF AI返却: %s",
            content,
        )

        result = parse_json_response(
            content
        )

        result["department"] = normalize_department(
            result["department"]
        )

        logger.info(
            "特調TB画像PDF 抽出結果: %s",
            result,
        )

        return result

    except Exception as error:
        logger.exception(
            "特調TB画像PDFのAI解析失敗: %s",
            path.name,
        )

        raise RuntimeError(
            "画像PDFのAI解析に失敗しました。\n\n"
            f"対象PDF：{path.name}\n"
            f"詳細：{error}"
        ) from error