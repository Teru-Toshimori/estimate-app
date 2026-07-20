from pathlib import Path

from PySide6.QtWidgets import (
    QApplication,
    QFileDialog,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QTabWidget,
    QVBoxLayout,
    QWidget,
)

from ui.tabs.msr_tab import MsrTab
from ui.tabs.tg_tab import TgTab
from ui.tabs.tokucho_other_tab import TokuchoOtherTab
from ui.tabs.tokucho_tb_tab import TokuchoTbTab


class MainWindow(QMainWindow):
    """
    見積書作成ツールのメイン画面。

    共通入力:
        ・業務計画書フォルダ
        ・出力フォルダ
        ・管理台帳URL
        ・利用者一覧URL

    タブ:
        ・特調TB
        ・特調TB以外
        ・TG
        ・MSR
    """

    def __init__(self):
        super().__init__()

        self.setWindowTitle(
            "見積書作成ツール"
        )

        self.resize(
            1200,
            850,
        )

        self.setup_ui()

    # =====================================
    # 画面作成
    # =====================================
    def setup_ui(self) -> None:

        central_widget = QWidget()

        self.setCentralWidget(
            central_widget
        )

        main_layout = QVBoxLayout(
            central_widget
        )

        main_layout.setContentsMargins(
            16,
            16,
            16,
            16,
        )

        main_layout.setSpacing(
            10
        )

        # =====================================
        # タイトル
        # =====================================
        title_label = QLabel(
            "見積書作成ツール"
        )

        title_label.setStyleSheet(
            "font-size: 22px;"
            "font-weight: bold;"
        )

        main_layout.addWidget(
            title_label
        )

        # =====================================
        # 共通入力タイトル
        # =====================================
        common_title_label = QLabel(
            "共通入力"
        )

        common_title_label.setStyleSheet(
            "font-size: 16px;"
            "font-weight: bold;"
            "margin-top: 4px;"
        )

        main_layout.addWidget(
            common_title_label
        )

        # =====================================
        # 共通入力エリア
        # =====================================
        common_input_layout = QHBoxLayout()

        common_input_layout.setSpacing(
            12
        )

        # =====================================
        # フォルダ設定グループ
        # =====================================
        folder_group = QGroupBox(
            "フォルダ設定"
        )

        folder_layout = QGridLayout(
            folder_group
        )

        folder_layout.setContentsMargins(
            12,
            14,
            12,
            12,
        )

        folder_layout.setHorizontalSpacing(
            8
        )

        folder_layout.setVerticalSpacing(
            10
        )

        # =====================================
        # 業務計画書フォルダ
        # =====================================
        pdf_folder_label = QLabel(
            "業務計画書フォルダ"
        )

        self.pdf_folder_edit = QLineEdit()

        self.pdf_folder_edit.setPlaceholderText(
            "PDFが入っているフォルダを選択"
        )

        self.pdf_folder_button = QPushButton(
            "参照"
        )

        self.pdf_folder_button.setFixedWidth(
            80
        )

        folder_layout.addWidget(
            pdf_folder_label,
            0,
            0,
        )

        folder_layout.addWidget(
            self.pdf_folder_edit,
            0,
            1,
        )

        folder_layout.addWidget(
            self.pdf_folder_button,
            0,
            2,
        )

        # =====================================
        # 出力フォルダ
        # =====================================
        output_folder_label = QLabel(
            "出力フォルダ"
        )

        self.output_folder_edit = QLineEdit()

        self.output_folder_edit.setPlaceholderText(
            "Excel・PDFの保存先を選択"
        )

        self.output_folder_button = QPushButton(
            "参照"
        )

        self.output_folder_button.setFixedWidth(
            80
        )

        folder_layout.addWidget(
            output_folder_label,
            1,
            0,
        )

        folder_layout.addWidget(
            self.output_folder_edit,
            1,
            1,
        )

        folder_layout.addWidget(
            self.output_folder_button,
            1,
            2,
        )

        folder_layout.setColumnStretch(
            1,
            1,
        )

        # =====================================
        # URL設定グループ
        # =====================================
        url_group = QGroupBox(
            "OneDrive／SharePoint設定"
        )

        url_layout = QGridLayout(
            url_group
        )

        url_layout.setContentsMargins(
            12,
            14,
            12,
            12,
        )

        url_layout.setHorizontalSpacing(
            8
        )

        url_layout.setVerticalSpacing(
            10
        )

        # =====================================
        # 管理台帳URL
        # =====================================
        ledger_url_label = QLabel(
            "管理台帳URL"
        )

        self.ledger_url_edit = QLineEdit()

        self.ledger_url_edit.setPlaceholderText(
            "管理台帳Excelの共有URLを入力"
        )

        self.ledger_url_button = QPushButton(
            "貼り付け"
        )

        self.ledger_url_button.setFixedWidth(
            80
        )

        url_layout.addWidget(
            ledger_url_label,
            0,
            0,
        )

        url_layout.addWidget(
            self.ledger_url_edit,
            0,
            1,
        )

        url_layout.addWidget(
            self.ledger_url_button,
            0,
            2,
        )

        # =====================================
        # 利用者一覧URL
        # =====================================
        user_master_url_label = QLabel(
            "利用者一覧URL"
        )

        self.user_master_url_edit = QLineEdit()

        self.user_master_url_edit.setPlaceholderText(
            "利用者一覧Excelの共有URLを入力"
        )

        self.user_master_url_button = QPushButton(
            "貼り付け"
        )

        self.user_master_url_button.setFixedWidth(
            80
        )

        url_layout.addWidget(
            user_master_url_label,
            1,
            0,
        )

        url_layout.addWidget(
            self.user_master_url_edit,
            1,
            1,
        )

        url_layout.addWidget(
            self.user_master_url_button,
            1,
            2,
        )

        url_layout.setColumnStretch(
            1,
            1,
        )

        # =====================================
        # グループを横並びに配置
        # =====================================
        common_input_layout.addWidget(
            folder_group,
            stretch=1,
        )

        common_input_layout.addWidget(
            url_group,
            stretch=1,
        )

        main_layout.addLayout(
            common_input_layout
        )

        # =====================================
        # タブ
        # =====================================
        self.tab_widget = QTabWidget()

        self.tokucho_tb_tab = TokuchoTbTab(
            input_provider=self.get_common_inputs
        )

        self.tokucho_other_tab = TokuchoOtherTab(
            input_provider=self.get_common_inputs
        )

        self.tg_tab = TgTab(
            input_provider=self.get_common_inputs
        )

        self.msr_tab = MsrTab(
            input_provider=self.get_common_inputs
        )

        self.tab_widget.addTab(
            self.tokucho_tb_tab,
            "特調TB",
        )

        self.tab_widget.addTab(
            self.tokucho_other_tab,
            "特調TB以外",
        )

        self.tab_widget.addTab(
            self.tg_tab,
            "TG",
        )

        self.tab_widget.addTab(
            self.msr_tab,
            "MSR",
        )

        main_layout.addWidget(
            self.tab_widget,
            stretch=1,
        )

        # =====================================
        # イベント接続
        # =====================================
        self.pdf_folder_button.clicked.connect(
            self.select_pdf_folder
        )

        self.output_folder_button.clicked.connect(
            self.select_output_folder
        )

        self.ledger_url_button.clicked.connect(
            self.paste_ledger_url
        )

        self.user_master_url_button.clicked.connect(
            self.paste_user_master_url
        )

    # =====================================
    # 共通入力取得
    # =====================================
    def get_common_inputs(self) -> dict:
        """
        各タブへ共通入力値を返す。
        """

        return {
            "pdf_folder": (
                self.pdf_folder_edit
                .text()
                .strip()
            ),
            "share_url": (
                self.ledger_url_edit
                .text()
                .strip()
            ),
            "user_master_url": (
                self.user_master_url_edit
                .text()
                .strip()
            ),
            "output_folder": (
                self.output_folder_edit
                .text()
                .strip()
            ),
        }

    # =====================================
    # 業務計画書フォルダ選択
    # =====================================
    def select_pdf_folder(self) -> None:

        current_path = (
            self.pdf_folder_edit
            .text()
            .strip()
        )

        if not Path(
            current_path
        ).is_dir():
            current_path = ""

        selected_folder = (
            QFileDialog.getExistingDirectory(
                self,
                "業務計画書フォルダを選択",
                current_path,
            )
        )

        if selected_folder:
            self.pdf_folder_edit.setText(
                selected_folder
            )

    # =====================================
    # 出力フォルダ選択
    # =====================================
    def select_output_folder(self) -> None:

        current_path = (
            self.output_folder_edit
            .text()
            .strip()
        )

        if not Path(
            current_path
        ).is_dir():
            current_path = ""

        selected_folder = (
            QFileDialog.getExistingDirectory(
                self,
                "出力フォルダを選択",
                current_path,
            )
        )

        if selected_folder:
            self.output_folder_edit.setText(
                selected_folder
            )

    # =====================================
    # 管理台帳URL貼り付け
    # =====================================
    def paste_ledger_url(self) -> None:

        self.paste_url_to_edit(
            target_edit=self.ledger_url_edit,
            field_name="管理台帳URL",
        )

    # =====================================
    # 利用者一覧URL貼り付け
    # =====================================
    def paste_user_master_url(self) -> None:

        self.paste_url_to_edit(
            target_edit=self.user_master_url_edit,
            field_name="利用者一覧URL",
        )

    # =====================================
    # URL共通貼り付け処理
    # =====================================
    def paste_url_to_edit(
        self,
        target_edit: QLineEdit,
        field_name: str,
    ) -> None:

        clipboard = (
            QApplication.clipboard()
        )

        url = (
            clipboard.text()
            .strip()
        )

        if not url:
            QMessageBox.warning(
                self,
                "確認",
                "クリップボードに"
                "URLがありません。",
            )
            return

        if not url.startswith(
            (
                "https://",
                "http://",
            )
        ):
            QMessageBox.warning(
                self,
                "確認",
                "クリップボードの内容が"
                f"{field_name}として使用できる"
                "URLではありません。",
            )
            return

        target_edit.setText(
            url
        )

    # =====================================
    # アプリ終了処理
    # =====================================
    def closeEvent(
        self,
        event,
    ) -> None:

        tabs = [
            self.tokucho_tb_tab,
            self.tokucho_other_tab,
            self.tg_tab,
            self.msr_tab,
        ]

        for tab in tabs:
            can_close_method = getattr(
                tab,
                "can_close",
                None,
            )

            if (
                callable(can_close_method)
                and not can_close_method()
            ):
                QMessageBox.warning(
                    self,
                    "処理中",
                    "現在処理を実行中のため、"
                    "アプリを終了できません。",
                )

                event.ignore()
                return

        event.accept()