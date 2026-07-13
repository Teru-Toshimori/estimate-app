from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QLabel,
    QVBoxLayout,
    QWidget,
)


class TokuchoOtherTab(QWidget):
    """特調TB以外の転記機能を配置するタブ。"""

    def __init__(self):
        super().__init__()

        self.setup_ui()

    def setup_ui(self):
        layout = QVBoxLayout(self)

        message = QLabel(
            "「特調TB以外」の機能は、今後ここへ実装します。"
        )

        message.setAlignment(Qt.AlignmentFlag.AlignCenter)

        layout.addWidget(message)