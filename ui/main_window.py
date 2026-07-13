from PySide6.QtWidgets import (
    QMainWindow,
    QMessageBox,
    QTabWidget,
)

from ui.tabs.msr_tab import MsrTab
from ui.tabs.tg_tab import TgTab
from ui.tabs.tokucho_other_tab import TokuchoOtherTab
from ui.tabs.tokucho_tb_tab import TokuchoTbTab


class MainWindow(QMainWindow):

    def __init__(self):
        super().__init__()

        self.setWindowTitle("見積書作成ツール")
        self.resize(1000, 820)

        self.setup_ui()

    def setup_ui(self):

        self.tab_widget = QTabWidget()
        self.setCentralWidget(self.tab_widget)

        self.tokucho_tb_tab = TokuchoTbTab()
        self.tokucho_other_tab = TokuchoOtherTab()
        self.tg_tab = TgTab()
        self.msr_tab = MsrTab()

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

        self.tab_widget.setCurrentIndex(0)

    def closeEvent(self, event):

        if not self.tokucho_tb_tab.can_close():
            QMessageBox.warning(
                self,
                "確認",
                "特調TBタブで台帳記入処理を実行中のため、"
                "アプリを終了できません。"
            )

            event.ignore()
            return

        event.accept()