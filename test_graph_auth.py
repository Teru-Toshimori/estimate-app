from services.graph_auth import GraphAuth


def main():
    try:
        auth = GraphAuth()
        token = auth.get_access_token()

        print()
        print("Microsoft Graph認証に成功しました。")
        print("アクセストークン取得:", bool(token))

    except Exception as error:
        print()
        print("Microsoft Graph認証に失敗しました。")
        print(error)


if __name__ == "__main__":
    main()