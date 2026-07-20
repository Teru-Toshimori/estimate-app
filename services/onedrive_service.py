import base64
import os
import tempfile
import time
from pathlib import Path
from typing import Any

import requests

from services.graph_auth import GraphAuth
from services.user_master_reader import (
    UserMasterReader,
)


class OneDriveService:
    """
    Microsoft Graphを使用して、
    OneDrive／SharePoint上のファイルを操作する。
    """

    GRAPH_BASE_URL = (
        "https://graph.microsoft.com/v1.0"
    )

    def __init__(
        self,
        device_flow_callback=None,
        user_master_url: str = "",
    ):
        self.auth = GraphAuth(
            device_flow_callback=(
                device_flow_callback
            )
        )

        # 利用者一覧URLは画面入力から受け取った値だけを使用する。
        self.user_master_url = str(
            user_master_url
        ).strip()

    # =====================================
    # 共有URLからDriveItem情報を取得
    # =====================================
    def resolve_shared_item(
        self,
        share_url: str,
    ) -> dict[str, Any]:

        share_url = share_url.strip()

        if not share_url:
            raise ValueError(
                "OneDrive／SharePointの共有URLが"
                "入力されていません。"
            )

        access_token = (
            self.auth.get_access_token()
        )

        share_token = self._encode_share_url(
            share_url
        )

        endpoint = (
            f"{self.GRAPH_BASE_URL}"
            f"/shares/{share_token}/driveItem"
        )

        response = requests.get(
            endpoint,
            headers={
                "Authorization": (
                    f"Bearer {access_token}"
                ),
                "Accept": "application/json",
                "Prefer": (
                    "redeemSharingLinkIfNecessary"
                ),
            },
            timeout=60,
        )

        self._raise_for_graph_error(
            response,
            "共有URLからファイル情報を"
            "取得できませんでした。",
        )

        item = response.json()

        remote_item = item.get(
            "remoteItem",
            {},
        )

        drive_id = (
            item.get(
                "parentReference",
                {},
            ).get("driveId")
            or remote_item.get(
                "parentReference",
                {},
            ).get("driveId")
        )

        item_id = (
            item.get("id")
            or remote_item.get("id")
        )

        file_name = (
            item.get("name")
            or remote_item.get("name")
        )

        etag = (
            item.get("eTag")
            or remote_item.get("eTag")
        )

        if not drive_id or not item_id:
            raise RuntimeError(
                "対象ファイルのdriveIdまたは"
                "itemIdを取得できませんでした。"
            )

        if not file_name:
            file_name = (
                "management_ledger.xlsx"
            )

        return {
            "drive_id": drive_id,
            "item_id": item_id,
            "file_name": file_name,
            "etag": etag,
        }

    # =====================================
    # Microsoftアカウント情報を取得
    # =====================================
    def get_current_user(
        self,
    ) -> dict[str, Any]:

        access_token = self.auth.get_access_token()

        endpoint = (
            f"{self.GRAPH_BASE_URL}"
            "/me"
            "?$select=displayName,mail,userPrincipalName,otherMails"
        )

        response = requests.get(
            endpoint,
            headers={
                "Authorization": f"Bearer {access_token}",
                "Accept": "application/json",
            },
            timeout=60,
        )

        self._raise_for_graph_error(
            response,
            "Microsoftアカウント情報を取得できませんでした。",
        )

        return response.json()

    # =====================================
    # Microsoftアカウントのメール取得
    # =====================================
    def get_current_user_email(
        self,
    ) -> str:
        """
        現在サインイン中のメールアドレスを取得する。

        優先順位：
        1. mail
        2. userPrincipalName
        """

        user = self.get_current_user()

        email_address = (
            user.get("mail")
            or user.get(
                "userPrincipalName"
            )
            or ""
        )

        return str(
            email_address
        ).strip()

    # =====================================
    # 利用者一覧から発行者名を取得
    # =====================================
    def get_issuer_name_from_master(self) -> str:
        """
        Graphから取得した複数のメール候補で
        利用者一覧Excelを検索する。
        """

        user = self.get_current_user()

        mail = str(
            user.get("mail") or ""
        ).strip()

        user_principal_name = str(
            user.get("userPrincipalName") or ""
        ).strip()

        display_name = str(
            user.get("displayName") or ""
        ).strip()

        other_mails = user.get("otherMails") or []

        email_candidates = [
            mail,
            user_principal_name,
            *[
                str(email).strip()
                for email in other_mails
                if str(email).strip()
            ],
        ]

        user_master_url = self.user_master_url

        if not user_master_url:
            raise ValueError(
                "利用者一覧URLが入力されていません。\n\n"
                "画面上部の共通入力欄にある"
                "「利用者一覧URL」へ、"
                "利用者一覧Excelの共有URLを"
                "入力してください。"
            )

        if not user_master_url.startswith(
            (
                "https://",
                "http://",
            )
        ):
            raise ValueError(
                "利用者一覧URLの形式が"
                "正しくありません。\n\n"
                f"{user_master_url}"
            )

        temp_path = None

        try:
            temp_path, _ = self.download_to_temp(
                user_master_url
            )

            reader = UserMasterReader()

            user_name = reader.find_user_name(
                excel_path=temp_path,
                email_addresses=email_candidates,
            )

            if user_name:
                return user_name

            raise ValueError(
                "利用者一覧Excelに、サインイン中のメールアドレスと"
                "一致する利用者が見つかりませんでした。\n\n"
                f"mail: {mail or 'なし'}\n"
                f"userPrincipalName: "
                f"{user_principal_name or 'なし'}"
            )

        finally:
            if temp_path and os.path.exists(temp_path):
                try:
                    os.remove(temp_path)
                except OSError:
                    pass

    # =====================================
    # 一時ファイルへダウンロード
    # =====================================
    def download_to_temp(
        self,
        share_url: str,
    ) -> tuple[
        str,
        dict[str, Any],
    ]:

        item = self.resolve_shared_item(
            share_url
        )

        access_token = (
            self.auth.get_access_token()
        )

        endpoint = (
            f"{self.GRAPH_BASE_URL}"
            f"/drives/{item['drive_id']}"
            f"/items/{item['item_id']}"
            "/content"
        )

        response = requests.get(
            endpoint,
            headers={
                "Authorization": (
                    f"Bearer {access_token}"
                ),
            },
            timeout=120,
            allow_redirects=True,
        )

        self._raise_for_graph_error(
            response,
            "OneDrive／SharePointから"
            "ファイルをダウンロード"
            "できませんでした。",
        )

        suffix = (
            Path(
                item["file_name"]
            ).suffix
            or ".xlsx"
        )

        file_descriptor, temp_path = (
            tempfile.mkstemp(
                prefix=(
                    "estimate_ledger_"
                ),
                suffix=suffix,
            )
        )

        os.close(file_descriptor)

        with open(
            temp_path,
            "wb",
        ) as temp_file:
            temp_file.write(
                response.content
            )

        return temp_path, item

    # =====================================
    # OneDrive／SharePointの元ファイルを上書き
    # =====================================
    def upload_replace(
        self,
        local_file_path: str,
        item: dict[str, Any],
    ) -> dict[str, Any]:

        local_path = Path(
            local_file_path
        )

        if not local_path.exists():
            raise FileNotFoundError(
                "アップロード対象ファイルが"
                "見つかりません。\n"
                f"{local_path}"
            )

        access_token = (
            self.auth.get_access_token()
        )

        endpoint = (
            f"{self.GRAPH_BASE_URL}"
            f"/drives/{item['drive_id']}"
            f"/items/{item['item_id']}"
            "/content"
        )

        headers = {
            "Authorization": (
                f"Bearer {access_token}"
            ),
            "Content-Type": (
                "application/vnd."
                "openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
        }

        if item.get("etag"):
            headers["If-Match"] = (
                item["etag"]
            )

        max_attempts = 5

        for attempt in range(
            1,
            max_attempts + 1,
        ):
            with local_path.open(
                "rb"
            ) as file_stream:
                response = requests.put(
                    endpoint,
                    headers=headers,
                    data=file_stream,
                    timeout=180,
                )

            if response.ok:
                return response.json()

            if response.status_code == 412:
                raise RuntimeError(
                    "管理台帳はダウンロード後に"
                    "別のユーザーによって"
                    "更新されました。\n"
                    "最新状態を取得してから"
                    "再実行してください。"
                )

            if response.status_code == 423:
                if attempt < max_attempts:
                    wait_seconds = (
                        attempt * 10
                    )

                    time.sleep(
                        wait_seconds
                    )

                    continue

                raise RuntimeError(
                    "管理台帳がロックされているため、"
                    "上書きできませんでした。\n\n"
                    "ブラウザー版Excel、"
                    "デスクトップ版Excel、"
                    "Teams、SharePointの"
                    "プレビュー画面で対象ファイルを"
                    "開いていないか確認してください。\n"
                    "他の利用者が編集中の場合は、"
                    "編集終了後に再実行してください。"
                )

            self._raise_for_graph_error(
                response,
                "OneDrive／SharePointへ"
                "ファイルを上書き"
                "できませんでした。",
            )

        raise RuntimeError(
            "OneDrive／SharePointへの"
            "上書きに失敗しました。"
        )

    # =====================================
    # ダウンロード → 台帳編集 → 再アップロード
    # =====================================
    def update_ledger(
        self,
        share_url: str,
        ledger_writer,
        data,
    ) -> dict[str, Any]:

        temp_path = None

        try:
            # 管理台帳をダウンロード
            temp_path, item = (
                self.download_to_temp(
                    share_url
                )
            )

            # 利用者一覧から発行者名を取得
            issuer_name = (
                self.get_issuer_name_from_master()
            )

            # 一時ファイルを編集
            ledger_result = (
                ledger_writer.write(
                    excel_path=temp_path,
                    data=data,
                    issuer_name=issuer_name,
                )
            )

            # 編集済みファイルを上書き
            uploaded_item = (
                self.upload_replace(
                    local_file_path=temp_path,
                    item=item,
                )
            )

            return {
                **ledger_result,
                "remote_name": (
                    uploaded_item.get(
                        "name",
                        item["file_name"],
                    )
                ),
                "web_url": (
                    uploaded_item.get(
                        "webUrl",
                        "",
                    )
                ),
                "issuer_name": (
                    issuer_name
                ),
            }

        finally:
            if (
                temp_path
                and os.path.exists(temp_path)
            ):
                try:
                    os.remove(temp_path)
                except OSError:
                    pass

    # =====================================
    # 共有URLをGraph API用トークンへ変換
    # =====================================
    def _encode_share_url(
        self,
        share_url: str,
    ) -> str:

        encoded = (
            base64.urlsafe_b64encode(
                share_url.encode(
                    "utf-8"
                )
            ).decode("ascii")
        )

        encoded = encoded.rstrip("=")

        return f"u!{encoded}"

    # =====================================
    # Graph APIエラー処理
    # =====================================
    def _raise_for_graph_error(
        self,
        response: requests.Response,
        message: str,
    ) -> None:

        if response.ok:
            return

        detail = ""

        try:
            payload = response.json()

            error = payload.get(
                "error",
                {},
            )

            code = error.get(
                "code",
                "",
            )

            graph_message = error.get(
                "message",
                "",
            )

            detail = (
                f"{code}: {graph_message}"
            ).strip(": ")

        except ValueError:
            detail = response.text[
                :500
            ]

        raise RuntimeError(
            f"{message}\n\n"
            f"HTTP {response.status_code}\n"
            f"{detail}"
        )