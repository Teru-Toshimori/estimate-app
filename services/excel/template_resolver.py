from pathlib import Path
import sys


class TemplateResolver:
    """
    resources配下のテンプレートを取得する共通クラス

    構成

    resources/
        tb/
            テンプレート.xlsx

        msr/
            テンプレート.xlsx
    """

    SUPPORTED_EXTENSIONS = {
        ".xlsx",
        ".xlsm",
        ".xltx",
        ".xltm",
    }

    @staticmethod
    def project_root() -> Path:
        """
        開発環境
        exe化
        の両方に対応
        """

        if getattr(sys, "frozen", False):
            return Path(sys.executable).resolve().parent

        return Path(__file__).resolve().parent.parent.parent

    @classmethod
    def resolve(
        cls,
        folder_name: str,
    ) -> Path:
        """
        resources内のテンプレートを取得する

        resolve("tb")

        resolve("msr")
        """

        template_dir = (
            cls.project_root()
            / "resources"
            / folder_name
        )

        if not template_dir.exists():
            raise FileNotFoundError(
                f"テンプレートフォルダが存在しません。\n\n{template_dir}"
            )

        files = []

        for file in template_dir.iterdir():

            if not file.is_file():
                continue

            if file.name.startswith("~$"):
                continue

            if file.suffix.lower() not in cls.SUPPORTED_EXTENSIONS:
                continue

            files.append(file)

        if len(files) == 0:
            raise FileNotFoundError(
                f"{template_dir} にテンプレートがありません。"
            )

        if len(files) >= 2:

            file_list = "\n".join(
                f.name for f in files
            )

            raise ValueError(
                "テンプレートが複数あります。\n\n"
                f"{file_list}"
            )

        return files[0]