import json
import sys
from pathlib import Path
from typing import Any


class AppConfig:
    """
    exeの外部に配置した設定ファイルを読み込む。
    """

    CONFIG_RELATIVE_PATH = (
        Path("settings")
        / "graph_config.json"
    )

    @classmethod
    def get_base_dir(cls) -> Path:
        """
        開発時：
            プロジェクトのルートフォルダ

        exe実行時：
            exeが配置されているフォルダ
        """

        if getattr(sys, "frozen", False):
            return Path(
                sys.executable
            ).resolve().parent

        return Path(
            __file__
        ).resolve().parent.parent

    @classmethod
    def get_config_path(cls) -> Path:
        """
        graph_config.jsonの絶対パスを返す。
        """

        return (
            cls.get_base_dir()
            / cls.CONFIG_RELATIVE_PATH
        )

    @classmethod
    def load_graph_config(
        cls,
    ) -> dict[str, Any]:
        """
        Microsoft Graph設定を読み込む。
        """

        config_path = cls.get_config_path()

        if not config_path.exists():
            raise FileNotFoundError(
                "Microsoft Graph設定ファイルが"
                "見つかりません。\n\n"
                f"{config_path}\n\n"
                "settingsフォルダ内に"
                "graph_config.jsonを配置してください。"
            )

        try:
            with config_path.open(
                "r",
                encoding="utf-8-sig",
            ) as config_file:
                config = json.load(config_file)

        except json.JSONDecodeError as error:
            raise ValueError(
                "graph_config.jsonの形式が"
                "正しくありません。\n\n"
                f"{error}"
            ) from error

        client_id = str(
            config.get(
                "client_id",
                "",
            )
        ).strip()

        tenant_id = str(
            config.get(
                "tenant_id",
                "",
            )
        ).strip()

        user_master_url = str(
            config.get(
                "user_master_url",
                "",
            )
        ).strip()

        scopes = config.get(
            "scopes",
            [
                "User.Read",
                "Files.ReadWrite",
            ],
        )

        if not client_id:
            raise ValueError(
                "graph_config.jsonに"
                "client_idが設定されていません。"
            )

        if not tenant_id:
            raise ValueError(
                "graph_config.jsonに"
                "tenant_idが設定されていません。"
            )

        if not user_master_url:
            raise ValueError(
                "graph_config.jsonに"
                "user_master_urlが設定されていません。"
            )

        if not user_master_url.startswith(
            ("https://", "http://")
        ):
            raise ValueError(
                "graph_config.jsonの"
                "user_master_urlがURL形式ではありません。"
            )

        if (
            not isinstance(scopes, list)
            or not scopes
        ):
            raise ValueError(
                "graph_config.jsonのscopesは、"
                "1件以上の配列で指定してください。"
            )

        normalized_scopes = [
            str(scope).strip()
            for scope in scopes
            if str(scope).strip()
        ]

        if not normalized_scopes:
            raise ValueError(
                "graph_config.jsonのscopesに"
                "有効な値がありません。"
            )

        return {
            "client_id": client_id,
            "tenant_id": tenant_id,
            "authority": (
                "https://login.microsoftonline.com/"
                f"{tenant_id}"
            ),
            "scopes": normalized_scopes,
            "user_master_url": user_master_url,
        }