import webbrowser

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QDialog,
    QLabel,
    QPushButton,
    QTextEdit,
    QVBoxLayout,
)


class DeviceLoginDialog(QDialog):
    """
    Microsoft Device Code認証ダイアログ
    """

    def __init__(
        self,
        flow: dict,
        parent=None,
    ):
        super().__init__(parent)

        self.flow = flow

        self.setWindowTitle(
            "Microsoftアカウント認証"
        )

        self.setModal(True)
        self.resize(650, 380)

        layout = QVBoxLayout(self)

        title = QLabel(
            "Microsoftアカウントへサインインしてください"
        )
        title.setStyleSheet(
            "font-size:18px;"
            "font-weight:bold;"
        )

        layout.addWidget(title)

        explanation = QLabel(
            "① 下のボタンから認証ページを開きます。\n"
            "② 表示されているコードを入力します。\n"
            "③ 認証完了後、この画面を閉じてください。"
        )

        explanation.setWordWrap(True)

        layout.addWidget(explanation)

        # 認証コード表示
        self.code_text = QTextEdit()
        self.code_text.setReadOnly(True)
        self.code_text.setMaximumHeight(60)

        user_code = self.flow.get(
            "user_code",
            ""
        )

        self.code_text.setPlainText(
            user_code
        )

        layout.addWidget(self.code_text)

        # 認証ページを開く
        open_button = QPushButton(
            "Microsoft認証ページを開く"
        )

        open_button.clicked.connect(
            self.open_login_page
        )

        layout.addWidget(open_button)

        # メッセージ全文
        message = self.flow.get(
            "message",
            "",
        )

        self.message_text = QTextEdit()

        self.message_text.setReadOnly(True)

        self.message_text.setPlainText(
            message
        )

        layout.addWidget(
            self.message_text
        )

        # 閉じる
        close_button = QPushButton(
            "認証完了"
        )

        close_button.clicked.connect(
            self.accept
        )

        layout.addWidget(
            close_button,
            alignment=Qt.AlignRight,
        )

    def open_login_page(self):
        """
        Device Code認証ページをブラウザで開く。

        Device Flowでは
        https://microsoft.com/devicelogin
        を開く。

        deviceauth や token を
        開いてしまうと
        AADSTS900561になるため防止する。
        """

        login_url = (
            self.flow.get(
                "verification_uri_complete"
            )
            or self.flow.get(
                "verification_uri"
            )
            or "https://microsoft.com/devicelogin"
        )

        # deviceauth や token が来た場合は強制的に正しいURLへ
        lower_url = login_url.lower()

        if (
            "deviceauth" in lower_url
            or "/token" in lower_url
        ):
            login_url = (
                "https://microsoft.com/devicelogin"
            )

        webbrowser.open(login_url)