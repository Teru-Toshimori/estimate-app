import webbrowser

from PySide6.QtWidgets import (
    QDialog,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QPushButton,
    QTextEdit,
    QVBoxLayout,
)


class DeviceLoginDialog(QDialog):

    def __init__(self, flow: dict, parent=None):
        super().__init__(parent)

        self.flow = flow

        self.setWindowTitle("Microsoftアカウント認証")
        self.resize(560, 360)
        self.setModal(True)

        self.setup_ui()

    def setup_ui(self):

        main_layout = QVBoxLayout(self)

        description = QLabel(
            "OneDrive／SharePointへ接続するため、"
            "Microsoftアカウントでサインインしてください。"
        )
        description.setWordWrap(True)

        main_layout.addWidget(description)

        url_label = QLabel("認証ページ")

        self.url_edit = QLineEdit()
        self.url_edit.setReadOnly(True)
        self.url_edit.setText(
            self.flow.get(
                "verification_uri",
                "https://microsoft.com/devicelogin",
            )
        )

        main_layout.addWidget(url_label)
        main_layout.addWidget(self.url_edit)

        code_label = QLabel("認証コード")

        self.code_edit = QLineEdit()
        self.code_edit.setReadOnly(True)
        self.code_edit.setText(
            self.flow.get("user_code", "")
        )

        main_layout.addWidget(code_label)
        main_layout.addWidget(self.code_edit)

        message_label = QLabel("認証案内")

        self.message_edit = QTextEdit()
        self.message_edit.setReadOnly(True)
        self.message_edit.setPlainText(
            self.flow.get("message", "")
        )

        main_layout.addWidget(message_label)
        main_layout.addWidget(self.message_edit)

        button_layout = QHBoxLayout()

        browser_button = QPushButton(
            "ブラウザーで認証ページを開く"
        )
        copy_button = QPushButton(
            "認証コードをコピー"
        )
        close_button = QPushButton(
            "認証を続行"
        )

        browser_button.clicked.connect(
            self.open_browser
        )
        copy_button.clicked.connect(
            self.copy_code
        )
        close_button.clicked.connect(
            self.accept
        )

        button_layout.addWidget(browser_button)
        button_layout.addWidget(copy_button)
        button_layout.addWidget(close_button)

        main_layout.addLayout(button_layout)

    def open_browser(self):

        url = self.url_edit.text().strip()

        if url:
            webbrowser.open(url)

    def copy_code(self):

        self.code_edit.selectAll()
        self.code_edit.copy()