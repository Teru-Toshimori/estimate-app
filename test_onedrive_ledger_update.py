from models.estimate_data import EstimateData
from services.ledger_writer import LedgerWriter
from services.onedrive_service import OneDriveService


def main():

    share_url = input(
        "テスト用管理台帳のOneDrive／SharePoint URLを入力してください：\n"
    ).strip()

    # 実際のPDF解析結果を模したテストデータ
    data = EstimateData(
        application_no="ITK99999",
        voucher_no="TEST-99999",
        department="第３シート設計部",
        subject="Graph API連携テスト",
        model_code="TEST",
        amount="100,000円",
        due_date="2026年08月31日",
        outputs=[
            "① テスト成果物",
        ],
    )

    confirm = input(
        "\n管理台帳へテスト行を追加し、元ファイルを上書きします。"
        "\n実行しますか？ [y/N]: "
    ).strip().lower()

    if confirm != "y":
        print("処理を中止しました。")
        return

    try:
        onedrive_service = OneDriveService()
        ledger_writer = LedgerWriter()

        result = onedrive_service.update_ledger(
            share_url=share_url,
            ledger_writer=ledger_writer,
            data=data,
        )

        print()
        print("=" * 60)
        print("台帳更新に成功しました")
        print("=" * 60)
        print("ファイル名:", result.get("remote_name", ""))
        print("シート名:", result["sheet_name"])
        print("追加行:", result["row"])
        print("見積番号:", result["estimate_no"])
        print("発行日:", result["issue_date"])
        print("発行者:", result["user_name"])
        print("Web URL:", result.get("web_url", ""))

    except Exception as error:
        print()
        print("=" * 60)
        print("台帳更新に失敗しました")
        print("=" * 60)
        print(error)


if __name__ == "__main__":
    main()