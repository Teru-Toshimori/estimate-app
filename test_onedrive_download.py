import os

from services.onedrive_service import OneDriveService


def main():
    share_url = input(
        "管理台帳のOneDrive／SharePoint URLを入力してください：\n"
    ).strip()

    service = OneDriveService()
    temp_path = None

    try:
        temp_path, item = service.download_to_temp(share_url)

        print()
        print("ダウンロードに成功しました。")
        print("元のファイル名:", item["file_name"])
        print("一時保存先:", temp_path)
        print("ファイルサイズ:", os.path.getsize(temp_path), "bytes")

        input(
            "\n一時ファイルを確認したらEnterを押してください。"
        )

    except Exception as error:
        print()
        print("ダウンロードに失敗しました。")
        print(error)

    finally:
        if temp_path and os.path.exists(temp_path):
            os.remove(temp_path)
            print("一時ファイルを削除しました。")


if __name__ == "__main__":
    main()