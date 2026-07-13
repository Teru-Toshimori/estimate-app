from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QLabel,
    QVBoxLayout,
    QWidget,
)


class TgTab(QWidget):
    """TG向けの転記機能を配置するタブ。"""

    def __init__(self):
        super().__init__()

        self.setup_ui()

    def setup_ui(self):
        layout = QVBoxLayout(self)

        message = QLabel(
            "「TG」の機能は、今後ここへ実装します。"
        )

        message.setAlignment(Qt.AlignmentFlag.AlignCenter)

        layout.addWidget(message)