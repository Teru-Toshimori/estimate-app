from enum import Enum
from typing import Mapping, Sequence

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QDialog,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QVBoxLayout,
    QWidget,
)


class DialogType(Enum):
    """ダイアログの表示種別。"""

    QUESTION = "question"
    SUCCESS = "success"
    WARNING = "warning"
    ERROR = "error"
    INFORMATION = "information"


class CustomMessageDialog(QDialog):
    """
    アプリ共通のメッセージダイアログ。

    実行確認、完了、中止、警告、エラーなどを
    同じレイアウトで表示する。
    """

    def __init__(
        self,
        title: str,
        heading: str,
        message: str = "",
        dialog_type: DialogType = DialogType.INFORMATION,
        confirm_text: str = "OK",
        cancel_text: str = "",
        parent: QWidget | None = None,
    ):
        super().__init__(parent)

        self.dialog_type = dialog_type
        self.confirm_text = confirm_text
        self.cancel_text = cancel_text

        self.setWindowTitle(title)
        self.setModal(True)
        self.setMinimumWidth(480)

        self.setup_ui(
            heading=heading,
            message=message,
        )

    # =====================================
    # 画面作成
    # =====================================
    def setup_ui(
        self,
        heading: str,
        message: str,
    ) -> None:

        self.main_layout = QVBoxLayout(self)
        self.main_layout.setContentsMargins(
            24,
            20,
            24,
            20,
        )
        self.main_layout.setSpacing(16)

        # 上部
        header_layout = QHBoxLayout()
        header_layout.setSpacing(14)

        icon_label = QLabel(
            self.get_icon_text()
        )
        icon_label.setAlignment(
            Qt.AlignmentFlag.AlignCenter
        )
        icon_label.setFixedSize(
            42,
            42,
        )
        icon_label.setStyleSheet(
            self.get_icon_style()
        )

        heading_label = QLabel(heading)
        heading_label.setWordWrap(True)
        heading_label.setStyleSheet(
            "font-size: 17px;"
            "font-weight: bold;"
        )

        header_layout.addWidget(icon_label)
        header_layout.addWidget(
            heading_label,
            stretch=1,
        )
        self.main_layout.addLayout(header_layout)

        # 区切り線
        separator = QFrame()
        separator.setFrameShape(
            QFrame.Shape.HLine
        )
        separator.setFrameShadow(
            QFrame.Shadow.Sunken
        )
        self.main_layout.addWidget(separator)

        # 本文・サマリーを追加する領域
        self.content_layout = QVBoxLayout()
        self.content_layout.setSpacing(12)
        self.main_layout.addLayout(
            self.content_layout
        )

        if message:
            self.add_message(message)

        self.main_layout.addStretch()

        # ボタン
        button_layout = QHBoxLayout()
        button_layout.addStretch()

        confirm_button = QPushButton(
            self.confirm_text
        )
        confirm_button.setMinimumWidth(110)
        confirm_button.setMinimumHeight(36)
        confirm_button.setDefault(True)
        confirm_button.clicked.connect(
            self.accept
        )
        button_layout.addWidget(
            confirm_button
        )

        if self.cancel_text:
            cancel_button = QPushButton(
                self.cancel_text
            )
            cancel_button.setMinimumWidth(110)
            cancel_button.setMinimumHeight(36)
            cancel_button.clicked.connect(
                self.reject
            )
            button_layout.addWidget(
                cancel_button
            )

        self.main_layout.addLayout(
            button_layout
        )

    # =====================================
    # 通常メッセージ
    # =====================================
    def add_message(
        self,
        message: str,
    ) -> None:

        message_label = QLabel(message)
        message_label.setWordWrap(True)
        message_label.setTextInteractionFlags(
            Qt.TextInteractionFlag.TextSelectableByMouse
        )
        message_label.setStyleSheet(
            "font-size: 14px;"
        )

        self.content_layout.addWidget(
            message_label
        )

    # =====================================
    # 件数サマリー
    # =====================================
    def add_summary(
        self,
        section_title: str,
        values: Mapping[str, int],
    ) -> None:

        if self.content_layout.count() > 0:
            separator = QFrame()
            separator.setFrameShape(
                QFrame.Shape.HLine
            )
            separator.setStyleSheet(
                "color: #D9D9D9;"
            )
            self.content_layout.addWidget(
                separator
            )

        title_label = QLabel(
            f"■ {section_title}"
        )
        title_label.setStyleSheet(
            "font-size: 14px;"
            "font-weight: bold;"
        )
        self.content_layout.addWidget(
            title_label
        )

        grid = QGridLayout()
        grid.setHorizontalSpacing(24)
        grid.setVerticalSpacing(8)
        grid.setColumnStretch(0, 1)

        grid_row = 0

        for label, count in values.items():
            label_text = str(label)

            name_label = QLabel(
                label_text
            )
            count_label = QLabel(
                f"{int(count)}件"
            )

            count_label.setAlignment(
                Qt.AlignmentFlag.AlignRight
                | Qt.AlignmentFlag.AlignVCenter
            )
            count_label.setMinimumWidth(70)
            count_label.setStyleSheet(
                self.get_summary_value_style(
                    label_text
                )
            )

            grid.addWidget(
                name_label,
                grid_row,
                0,
            )
            grid.addWidget(
                count_label,
                grid_row,
                1,
            )

            grid_row += 1

            # 「処理済み」と結果内訳の間に区切り線を表示する。
            if label_text == "処理済み":
                result_separator = QFrame()
                result_separator.setFrameShape(
                    QFrame.Shape.HLine
                )
                result_separator.setFrameShadow(
                    QFrame.Shadow.Plain
                )
                result_separator.setStyleSheet(
                    "color: #D9D9D9;"
                    "margin-top: 3px;"
                    "margin-bottom: 3px;"
                )

                grid.addWidget(
                    result_separator,
                    grid_row,
                    0,
                    1,
                    2,
                )

                grid_row += 1

        self.content_layout.addLayout(grid)

    def get_summary_value_style(
        self,
        label: str,
    ) -> str:

        if label == "成功":
            color = "#2E7D32"
        elif label in {
            "対象外",
            "NG",
            "中止",
        }:
            color = "#ED6C02"
        elif label == "失敗":
            color = "#D32F2F"
        else:
            color = "#333333"

        return (
            f"color: {color};"
            "font-weight: bold;"
        )

    # =====================================
    # アイコン
    # =====================================
    def get_icon_text(self) -> str:

        icons = {
            DialogType.QUESTION: "?",
            DialogType.SUCCESS: "✓",
            DialogType.WARNING: "!",
            DialogType.ERROR: "×",
            DialogType.INFORMATION: "i",
        }

        return icons.get(
            self.dialog_type,
            "i",
        )

    def get_icon_style(self) -> str:

        background_colors = {
            DialogType.QUESTION: "#1976D2",
            DialogType.SUCCESS: "#2E7D32",
            DialogType.WARNING: "#ED6C02",
            DialogType.ERROR: "#D32F2F",
            DialogType.INFORMATION: "#1976D2",
        }

        background_color = (
            background_colors.get(
                self.dialog_type,
                "#1976D2",
            )
        )

        return (
            f"background-color: {background_color};"
            "color: white;"
            "border-radius: 21px;"
            "font-size: 22px;"
            "font-weight: bold;"
        )

    # =====================================
    # 実行確認
    # =====================================
    @classmethod
    def confirm(
        cls,
        parent: QWidget | None,
        title: str,
        heading: str,
        message: str,
        confirm_text: str = "実行",
        cancel_text: str = "キャンセル",
    ) -> bool:

        dialog = cls(
            parent=parent,
            title=title,
            heading=heading,
            message=message,
            dialog_type=DialogType.QUESTION,
            confirm_text=confirm_text,
            cancel_text=cancel_text,
        )

        return (
            dialog.exec()
            == QDialog.DialogCode.Accepted
        )

    # =====================================
    # 完了サマリー
    # =====================================
    @classmethod
    def summary(
        cls,
        parent: QWidget | None,
        title: str,
        heading: str,
        sections: Sequence[
            tuple[str, Mapping[str, int]]
        ],
        dialog_type: DialogType = DialogType.SUCCESS,
    ) -> None:

        dialog = cls(
            parent=parent,
            title=title,
            heading=heading,
            dialog_type=dialog_type,
            confirm_text="OK",
        )

        for section_title, values in sections:
            dialog.add_summary(
                section_title,
                values,
            )

        dialog.exec()

    # =====================================
    # 完了
    # =====================================
    @classmethod
    def success(
        cls,
        parent: QWidget | None,
        title: str,
        heading: str,
        message: str,
    ) -> None:

        dialog = cls(
            parent=parent,
            title=title,
            heading=heading,
            message=message,
            dialog_type=DialogType.SUCCESS,
            confirm_text="OK",
        )
        dialog.exec()

    # =====================================
    # 情報
    # =====================================
    @classmethod
    def information(
        cls,
        parent: QWidget | None,
        title: str,
        heading: str,
        message: str,
    ) -> None:

        dialog = cls(
            parent=parent,
            title=title,
            heading=heading,
            message=message,
            dialog_type=DialogType.INFORMATION,
            confirm_text="OK",
        )
        dialog.exec()

    # =====================================
    # 警告
    # =====================================
    @classmethod
    def warning(
        cls,
        parent: QWidget | None,
        title: str,
        heading: str,
        message: str,
    ) -> None:

        dialog = cls(
            parent=parent,
            title=title,
            heading=heading,
            message=message,
            dialog_type=DialogType.WARNING,
            confirm_text="OK",
        )
        dialog.exec()

    # =====================================
    # エラー
    # =====================================
    @classmethod
    def error(
        cls,
        parent: QWidget | None,
        title: str,
        heading: str,
        message: str,
    ) -> None:

        dialog = cls(
            parent=parent,
            title=title,
            heading=heading,
            message=message,
            dialog_type=DialogType.ERROR,
            confirm_text="OK",
        )
        dialog.exec()
