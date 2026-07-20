from collections.abc import Callable
from pathlib import Path
import os

import msal

from services.app_config import AppConfig


class GraphAuth:
    """Microsoft Graph用アクセストークンを取得する。"""

    def __init__(
        self,
        device_flow_callback: Callable[[dict], None] | None = None,
    ):
        self.device_flow_callback = device_flow_callback

        config = AppConfig.load_graph_config()

        self.client_id = config["client_id"]
        self.tenant_id = config["tenant_id"]
        self.scopes = config["scopes"]

        if not self.client_id:
            raise ValueError(
                "client_idが設定されていません。"
            )

        if not self.tenant_id:
            raise ValueError(
                "tenant_idが設定されていません。"
            )

        if not self.scopes:
            raise ValueError(
                "scopesが設定されていません。"
            )

        self.authority = (
            "https://login.microsoftonline.com/"
            f"{self.tenant_id}"
        )

        self.cache_path = self._get_cache_path()
        self.cache = msal.SerializableTokenCache()

        if self.cache_path.exists():
            try:
                cache_text = self.cache_path.read_text(
                    encoding="utf-8"
                )
                self.cache.deserialize(cache_text)

            except Exception:
                pass

        self.app = msal.PublicClientApplication(
            client_id=self.client_id,
            authority=self.authority,
            token_cache=self.cache,
        )

    def get_access_token(self) -> str:
        result = self._acquire_token_silent()

        if not result:
            result = self._acquire_token_by_device_flow()

        self._save_cache()

        access_token = result.get("access_token")

        if not access_token:
            error = (
                result.get("error_description")
                or result.get("error")
                or "不明な認証エラー"
            )

            raise RuntimeError(
                "Microsoftアカウントの認証に失敗しました。\n\n"
                f"{error}"
            )

        return access_token

    def _acquire_token_silent(self):
        accounts = self.app.get_accounts()

        if not accounts:
            return None

        return self.app.acquire_token_silent(
            scopes=self.scopes,
            account=accounts[0],
        )

    def _acquire_token_by_device_flow(self):
        flow = self.app.initiate_device_flow(
            scopes=self.scopes,
        )

        if "user_code" not in flow:
            raise RuntimeError(
                "デバイスコード認証を開始できませんでした。\n\n"
                f"{flow}"
            )

        # ブラウザで開くURLを明示的に設定
        flow["verification_uri"] = (
            flow.get("verification_uri")
            or "https://microsoft.com/devicelogin"
        )

        if (
            "deviceauth"
            in flow["verification_uri"].lower()
        ):
            flow["verification_uri"] = (
                "https://microsoft.com/devicelogin"
            )

        if self.device_flow_callback:
            self.device_flow_callback(flow)
        else:
            print(flow.get("message", ""))

        return (
            self.app.acquire_token_by_device_flow(
                flow
            )
        )

    def _save_cache(self):
        if self.cache.has_state_changed:
            self.cache_path.parent.mkdir(
                parents=True,
                exist_ok=True,
            )

            self.cache_path.write_text(
                self.cache.serialize(),
                encoding="utf-8",
            )

    def _get_cache_path(self) -> Path:
        """
        認証キャッシュはLocalAppDataへ保存する。

        CLIENT_ID・TENANT_IDごとにファイルを分けることで、
        設定変更後に別アカウントのキャッシュが混ざるのを防ぐ。
        """

        local_app_data = os.getenv("LOCALAPPDATA")

        if local_app_data:
            cache_dir = (
                Path(local_app_data)
                / "EstimateTool"
            )
        else:
            cache_dir = (
                Path.home()
                / ".estimate_tool"
            )

        safe_tenant_id = self.tenant_id.replace("-", "")
        safe_client_id = self.client_id.replace("-", "")

        cache_name = (
            f"token_cache_"
            f"{safe_tenant_id}_"
            f"{safe_client_id}.bin"
        )

        return cache_dir / cache_name