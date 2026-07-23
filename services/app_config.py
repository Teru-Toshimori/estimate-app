import os
import sys
from pathlib import Path
from typing import Any

from dotenv import load_dotenv


class AppConfig:
    """
    アプリ全体の設定を.envから読み込むクラス。

    開発環境:
        プロジェクトルートの.envを使用する。

    exe実行時:
        exeファイルと同じフォルダの.envを使用する。
    """

    ENV_FILE_NAME = ".env"

    _loaded = False

    # =====================================
    # 基準フォルダ
    # =====================================
    @classmethod
    def get_base_dir(cls) -> Path:
        """
        設定ファイルを検索する基準フォルダを返す。

        通常実行時:
            servicesフォルダの1つ上を
            プロジェクトルートとして使用する。

        exe実行時:
            exeファイルが存在するフォルダを使用する。
        """

        if getattr(sys, "frozen", False):
            return Path(
                sys.executable
            ).resolve().parent

        return (
            Path(__file__)
            .resolve()
            .parent
            .parent
        )

    # =====================================
    # .envのパス
    # =====================================
    @classmethod
    def get_env_path(cls) -> Path:
        """
        .envファイルの絶対パスを返す。
        """

        return (
            cls.get_base_dir()
            / cls.ENV_FILE_NAME
        )

    # =====================================
    # .envの読み込み
    # =====================================
    @classmethod
    def load_env(cls) -> Path:
        """
        .envファイルを読み込む。

        すでに読み込み済みの場合は、
        再読み込みせずにパスだけを返す。
        """

        env_path = cls.get_env_path()

        if not env_path.exists():
            raise FileNotFoundError(
                "設定ファイル「.env」が"
                "見つかりません。\n\n"
                f"確認先：{env_path}\n\n"
                "プロジェクトルート、または"
                "exeと同じフォルダに"
                ".envを配置してください。"
            )

        if not cls._loaded:
            load_dotenv(
                dotenv_path=env_path,
                override=False,
            )

            cls._loaded = True

        return env_path

    # =====================================
    # 必須設定値
    # =====================================
    @classmethod
    def get_required_value(
        cls,
        key: str,
        display_name: str | None = None,
    ) -> str:
        """
        .envから必須設定値を取得する。

        値が存在しない場合や空の場合は、
        ValueErrorを発生させる。
        """

        cls.load_env()

        value = str(
            os.getenv(
                key,
                "",
            )
        ).strip()

        if not value:
            raise ValueError(
                ".envに必要な設定が"
                "ありません。\n\n"
                f"設定名：{display_name or key}"
            )

        return value

    # =====================================
    # 任意設定値
    # =====================================
    @classmethod
    def get_optional_value(
        cls,
        key: str,
        default: str = "",
    ) -> str:
        """
        .envから任意設定値を取得する。

        設定されていない場合は、
        defaultで指定した値を返す。
        """

        cls.load_env()

        return str(
            os.getenv(
                key,
                default,
            )
        ).strip()

    # =====================================
    # OpenAI APIキー
    # =====================================
    @classmethod
    def get_openai_api_key(cls) -> str:
        """
        OpenAI APIキーを取得する。
        """

        return cls.get_required_value(
            key="OPENAI_API_KEY",
            display_name="OpenAI API Key",
        )

    # =====================================
    # Microsoft Graph Client ID
    # =====================================
    @classmethod
    def get_graph_client_id(cls) -> str:
        """
        Microsoft Graphの
        Client IDを取得する。
        """

        return cls.get_required_value(
            key="GRAPH_CLIENT_ID",
            display_name=(
                "Microsoft Graph Client ID"
            ),
        )

    # =====================================
    # Microsoft Graph Tenant ID
    # =====================================
    @classmethod
    def get_graph_tenant_id(cls) -> str:
        """
        Microsoft Graphの
        Tenant IDを取得する。
        """

        return cls.get_required_value(
            key="GRAPH_TENANT_ID",
            display_name=(
                "Microsoft Graph Tenant ID"
            ),
        )

    # =====================================
    # Microsoft Graph Scopes
    # =====================================
    @classmethod
    def get_graph_scopes(cls) -> list[str]:
        """
        Microsoft Graphの権限一覧を取得する。

        .envではカンマ区切りで設定する。

        例:
            GRAPH_SCOPES=
            User.Read,Files.ReadWrite,
            Files.ReadWrite.All
        """

        scopes_text = cls.get_optional_value(
            key="GRAPH_SCOPES",
            default=(
                "User.Read,"
                "Files.ReadWrite,"
                "Files.ReadWrite.All"
            ),
        )

        scopes = [
            scope.strip()
            for scope in scopes_text.split(",")
            if scope.strip()
        ]

        if not scopes:
            raise ValueError(
                ".envのGRAPH_SCOPESに"
                "有効な権限が"
                "設定されていません。"
            )

        return scopes

    # =====================================
    # Microsoft Graph Authority
    # =====================================
    @classmethod
    def get_graph_authority(cls) -> str:
        """
        Microsoft Graph認証用の
        Authority URLを返す。
        """

        tenant_id = (
            cls.get_graph_tenant_id()
        )

        return (
            "https://login.microsoftonline.com/"
            f"{tenant_id}"
        )

    # =====================================
    # Microsoft Graph設定一式
    # =====================================
    @classmethod
    def load_graph_config(
        cls,
    ) -> dict[str, Any]:
        """
        Microsoft Graph設定を
        .envからまとめて取得する。

        graph_config.jsonで使用していた形式と
        同じキー構成で返す。
        """

        client_id = (
            cls.get_graph_client_id()
        )

        tenant_id = (
            cls.get_graph_tenant_id()
        )

        scopes = (
            cls.get_graph_scopes()
        )

        authority = (
            "https://login.microsoftonline.com/"
            f"{tenant_id}"
        )

        return {
            "client_id": client_id,
            "tenant_id": tenant_id,
            "authority": authority,
            "scopes": scopes,
        }