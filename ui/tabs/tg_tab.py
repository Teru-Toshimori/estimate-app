from collections.abc import Callable
from pathlib import Path

from PySide6.QtWidgets import (
    QLabel,
    QMessageBox,
    QPushButton,
    QVBoxLayout,
    QWidget,
)


class TgTab(QWidget):
    """
    TG向けタブ。

    PDFフォルダ、管理台帳URL、出力フォルダは、
    MainWindowに配置した共通入力欄から取得する。

    現段階では、共通入力を正しく取得できるか
    確認する機能まで実装する。
    """

    def __init__(
        self,
        input_provider: Callable[[], dict],
        parent=None,
    ):
        super().__init__(parent)

        # MainWindowの共通入力取得メソッド
        self.input_provider = input_provider

        self.setup_ui()

    # =====================================
    # 画面作成
    # =====================================
    def setup_ui(self) -> None:

        main_layout = QVBoxLayout(self)

        main_layout.setContentsMargins(
            12,
            12,
            12,
            12,
        )

        main_layout.setSpacing(8)

        # =====================================
        # タイトル
        # =====================================
        title_label = QLabel(
            "TG"
        )

        title_label.setStyleSheet(
            "font-size: 18px;"
            "font-weight: bold;"
        )

        main_layout.addWidget(
            title_label
        )

        # =====================================
        # 説明
        # =====================================
        description_label = QLabel(
            "画面上部の共通入力欄に指定された"
            "業務計画書フォルダ、管理台帳URL、"
            "出力フォルダを使用します。\n\n"
            "TG向けの個別処理は、"
            "次の工程で実装します。"
        )

        description_label.setWordWrap(
            True
        )

        main_layout.addWidget(
            description_label
        )

        # =====================================
        # 共通入力確認ボタン
        # =====================================
        self.confirm_button = QPushButton(
            "共通入力を確認"
        )

        self.confirm_button.setMinimumHeight(
            42
        )

        self.confirm_button.clicked.connect(
            self.confirm_common_inputs
        )

        main_layout.addWidget(
            self.confirm_button
        )

        main_layout.addStretch()

    # =====================================
    # 共通入力確認
    # =====================================
    def confirm_common_inputs(self) -> None:

        inputs = self.input_provider()

        pdf_folder = str(
            inputs.get(
                "pdf_folder",
                "",
            )
        ).strip()

        share_url = str(
            inputs.get(
                "share_url",
                "",
            )
        ).strip()

        output_folder = str(
            inputs.get(
                "output_folder",
                "",
            )
        ).strip()

        # =====================================
        # 入力チェック
        # =====================================
        if not pdf_folder:
            QMessageBox.warning(
                self,
                "確認",
                "画面上部で業務計画書フォルダを"
                "選択してください。",
            )
            return

        pdf_folder_path = Path(
            pdf_folder
        )

        if not pdf_folder_path.exists():
            QMessageBox.warning(
                self,
                "確認",
                "指定された業務計画書フォルダが"
                "見つかりません。\n\n"
                f"{pdf_folder}",
            )
            return

        if not pdf_folder_path.is_dir():
            QMessageBox.warning(
                self,
                "確認",
                "業務計画書フォルダには"
                "フォルダを指定してください。",
            )
            return

        if not share_url:
            QMessageBox.warning(
                self,
                "確認",
                "画面上部で管理台帳URLを"
                "入力してください。",
            )
            return

        if not share_url.startswith(
            (
                "https://",
                "http://",
            )
        ):
            QMessageBox.warning(
                self,
                "確認",
                "管理台帳URLの形式が"
                "正しくありません。",
            )
            return

        if not output_folder:
            QMessageBox.warning(
                self,
                "確認",
                "画面上部で出力フォルダを"
                "選択してください。",
            )
            return

        output_folder_path = Path(
            output_folder
        )

        if not output_folder_path.exists():
            QMessageBox.warning(
                self,
                "確認",
                "指定された出力フォルダが"
                "見つかりません。\n\n"
                f"{output_folder}",
            )
            return

        if not output_folder_path.is_dir():
            QMessageBox.warning(
                self,
                "確認",
                "出力フォルダには"
                "フォルダを指定してください。",
            )
            return

        # =====================================
        # PDF件数取得
        # =====================================
        pdf_files = [
            path
            for path in pdf_folder_path.iterdir()
            if (
                path.is_file()
                and path.suffix.lower()
                == ".pdf"
            )
        ]

        pdf_count = len(
            pdf_files
        )

        # =====================================
        # 確認表示
        # =====================================
        QMessageBox.information(
            self,
            "共通入力確認",
            "TGタブで共通入力を取得しました。\n\n"
            f"業務計画書フォルダ：\n"
            f"{pdf_folder}\n\n"
            f"PDF件数：{pdf_count}件\n\n"
            f"管理台帳URL：\n"
            f"{share_url}\n\n"
            f"出力フォルダ：\n"
            f"{output_folder}",
        )

    # =====================================
    # アプリ終了可能判定
    # =====================================
    def can_close(self) -> bool:
        return True