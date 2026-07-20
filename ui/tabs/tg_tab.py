import glob
import logging
import os

from PySide6.QtWidgets import (
    QApplication,
    QFileDialog,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QTextEdit,
    QVBoxLayout,
    QWidget,
    QProgressBar,
)

from services.tg_excel_writer import TgExcelWriter
from services.tg_pdf_reader import TgPdfReader

from datetime import datetime

from services.tg_ledger_writer import TgLedgerWriter


logger = logging.getLogger(__name__)


class TgTab(QWidget):

    # =====================================
    # コンストラクタ
    # =====================================
    # TGタブ生成時に一度だけ呼ばれる
    #
    # ・PDF解析結果を保持するリスト初期化
    # ・画面(UI)作成
    #
    def __init__(self, parent=None):

        super().__init__(parent)

        # =====================================
        # PDF解析結果保持
        # =====================================
        # 複数PDF分保持
        #
        # [
        #   {
        #       "pdf_path": "xxx.pdf",
        #       "data": {
        #           "品名": "",
        #           "金額": ""
        #       }
        #   }
        # ]
        #
        self.pdf_results = []

        self.setup_ui()

    """
    TGタブの画面を作成する

    配置するもの
    ・入力フォルダ
    ・出力フォルダ
    ・操作ボタン
    ・進捗バー
    ・ログ
    ・状態表示

    最後に各ボタンのイベントを登録する
    """
    def setup_ui(self):

        layout = QVBoxLayout(self)

        # =====================================
        # タイトル
        # =====================================
        title = QLabel("TG 見積書作成")
        title.setStyleSheet(
            "font-size:18px;font-weight:bold;"
        )
        layout.addWidget(title)

        # =====================================
        # 入力フォルダ
        # =====================================
        input_label = QLabel("入力PDFフォルダ")

        self.input_dir_edit = QLineEdit()
        self.input_dir_edit.setPlaceholderText(
            "PDFが保存されているフォルダを選択してください"
        )

        self.input_dir_button = QPushButton("参照")

        input_layout = QHBoxLayout()
        input_layout.addWidget(self.input_dir_edit)
        input_layout.addWidget(self.input_dir_button)

        layout.addWidget(input_label)
        layout.addLayout(input_layout)

        # =====================================
        # 出力フォルダ
        # =====================================
        output_label = QLabel("出力フォルダ")

        self.output_dir_edit = QLineEdit()
        self.output_dir_edit.setPlaceholderText(
            "Excel・PDFの出力先フォルダ"
        )

        self.output_dir_button = QPushButton("参照")

        output_layout = QHBoxLayout()
        output_layout.addWidget(self.output_dir_edit)
        output_layout.addWidget(self.output_dir_button)

        layout.addWidget(output_label)
        layout.addLayout(output_layout)

        # =====================================
        # 操作ボタン
        # =====================================
        self.pdf_parse_button = QPushButton("PDF解析")
        self.ledger_write_button = QPushButton("台帳記入")
        self.output_button = QPushButton("Excel・PDF出力")

        button_layout = QHBoxLayout()
        button_layout.addWidget(self.pdf_parse_button)
        button_layout.addWidget(self.ledger_write_button)
        button_layout.addWidget(self.output_button)

        layout.addLayout(button_layout)

        # =====================================
        # 進捗バー
        # =====================================
        self.progress = QProgressBar()
        self.progress.setMinimum(0)
        self.progress.setValue(0)

        layout.addWidget(self.progress)

        # =====================================
        # ログ表示
        # =====================================
        log_label = QLabel("実行ログ")

        self.log_edit = QTextEdit()
        self.log_edit.setReadOnly(True)
        self.log_edit.setMinimumHeight(220)

        layout.addWidget(log_label)
        layout.addWidget(self.log_edit)

        # =====================================
        # 状態表示
        # =====================================
        self.status_label = QLabel("待機中")
        self.status_label.setStyleSheet(
            "color:#555;"
        )

        layout.addWidget(self.status_label)

        # =====================================
        # イベント
        # =====================================
        self.input_dir_button.clicked.connect(
            self.select_input_folder
        )

        self.output_dir_button.clicked.connect(
            self.select_output_folder
        )

        self.pdf_parse_button.clicked.connect(
            self.pdf_parse
        )

        self.ledger_write_button.clicked.connect(
            self.write_ledger
        )

        self.output_button.clicked.connect(
            self.export_excel_pdf
        )

    """
    画面の状態表示更新

    ・ステータスラベル
    ・ログ
    ・ProgressBar

    をまとめて更新する
    """
    def set_status(
        self,
        message: str,
        progress: int | None = None
    ):
        """
        状態表示更新
        """

        # ステータス表示
        self.status_label.setText(message)

        # ログへ出力
        self.add_log(message)

        # ProgressBar更新
        if progress is not None:
            self.progress.setValue(progress)

        # 画面更新
        QApplication.processEvents()

    """
    画面下のログへ1行追加する

    ・現在時刻付与
    ・最終行まで自動スクロール
    """
    def add_log(
        self,
        message: str
    ):
        """
        ログ表示
        """

        now = datetime.now().strftime(
            "%H:%M:%S"
        )

        self.log_edit.append(
            f"[{now}] {message}"
        )

        # 一番下へスクロール
        scrollbar = self.log_edit.verticalScrollBar()
        scrollbar.setValue(
            scrollbar.maximum()
        )

        # 画面更新
        QApplication.processEvents()

    """
    ログ画面を初期化する
    """
    def clear_log(self):

        self.log_text.clear()

        QApplication.processEvents()

    """
    入力フォルダ選択

    選択後

    ・パス表示
    ・PDF一覧取得
    ・ログ表示

    を行う
    """
    def select_input_folder(self):
        """
        入力PDFフォルダ選択
        """

        folder = QFileDialog.getExistingDirectory(
            self,
            "PDFフォルダを選択"
        )

        if not folder:
            return

        self.input_dir_edit.setText(folder)

        # ログ表示
        self.add_log("========== 入力フォルダ ==========")
        self.add_log(folder)

        # PDF一覧表示
        pdf_files = sorted(
            glob.glob(
                os.path.join(folder, "*.pdf")
            )
        )

        if pdf_files:

            self.add_log("")
            self.add_log("読込対象PDF")

            for pdf in pdf_files:

                self.add_log(
                    f"・{os.path.basename(pdf)}"
                )

            self.add_log(
                f"合計 {len(pdf_files)} 件"
            )

        else:

            self.add_log("")
            self.add_log("PDFが見つかりません。")

    """
    出力フォルダ選択

    選択後

    ・パス表示
    ・既存Excel
    ・既存PDF

    をログへ表示する
    """
    def select_output_folder(self):
        """
        出力フォルダ選択
        """

        folder = QFileDialog.getExistingDirectory(
            self,
            "出力フォルダを選択"
        )

        if not folder:
            return

        self.output_dir_edit.setText(folder)

        self.add_log("")
        self.add_log("========== 出力フォルダ ==========")
        self.add_log(folder)

        # 既存の出力ファイルがあれば表示
        excel_files = sorted(
            glob.glob(
                os.path.join(folder, "*.xlsm")
            )
        )

        pdf_files = sorted(
            glob.glob(
                os.path.join(folder, "*.pdf")
            )
        )

        if excel_files or pdf_files:

            self.add_log("")
            self.add_log("既存ファイル")

            for file in excel_files:
                self.add_log(
                    f"Excel : {os.path.basename(file)}"
                )

            for file in pdf_files:
                self.add_log(
                    f"PDF   : {os.path.basename(file)}"
                )

        else:

            self.add_log("")
            self.add_log("出力先フォルダは空です。")


    """
    入力フォルダ内のPDFを一括解析する

    処理

    入力フォルダ取得
        ↓
    PDF一覧取得
        ↓
    Excel一覧取得
        ↓
    PDFとExcelをペアリング
        ↓
    PDF解析(OpenAI Vision)
        ↓
    OCR結果保持
    """
    def pdf_parse(self):

        try:

            logger.info(
                "TG PDF解析ボタン押下"
            )

            # =====================================
            # 前回解析結果クリア
            # =====================================
            self.pdf_results.clear()


            # -----------------------------
            # 入力フォルダ取得
            # -----------------------------

            input_dir = self.input_dir_edit.text().strip()


            if not input_dir:

                QMessageBox.warning(
                    self,
                    "確認",
                    "入力PDFフォルダを選択してください"
                )

                return


            # 入力フォルダ内のPDF一覧取得
            pdf_files = sorted(
                glob.glob(
                    os.path.join(
                        input_dir,
                        "*.pdf"
                    )
                )
            )


            if not pdf_files:

                QMessageBox.warning(
                    self,
                    "確認",
                    "PDFがありません"
                )

                return

            logger.info(
                "解析PDF数 : %s",
                len(pdf_files)
            )

            # 入力フォルダ内のExcel一覧取得
            # PDFとペアになるExcelを探すため
            excel_files = sorted(
                glob.glob(
                    os.path.join(
                        input_dir,
                        "*.xlsm"
                    )
                )
            )

            logger.info(
                "Excel数 : %d",
                len(excel_files)
            )

            # PDF解析クラス生成
            reader = TgPdfReader()

            # -----------------------------
            # PDF解析
            # -----------------------------

            # PDFを1件ずつ解析
            for pdf_path in pdf_files:
                
                # PDFに対応するExcelを検索
                excel_path = self.find_excel_for_pdf(
                    pdf_path,
                    excel_files
                )

                if excel_path is None:

                    logger.warning(
                        "Excelが見つからないためスキップ : %s",
                        pdf_path
                    )

                    continue

                logger.info(
                    "解析開始 : %s",
                    pdf_path
                )


                result = reader.parse(
                    pdf_path
                )

                # =====================================
                # OCR結果保持
                #
                # {
                #   pdf_path
                #   excel_path
                #   data
                # }
                #
                # 後続の
                # ・台帳記入
                # ・Excel出力
                #
                # で利用する
                # =====================================
                self.pdf_results.append({
                    "pdf_path": pdf_path,
                    "excel_path": excel_path,
                    "data": result
                })

                logger.info("========== PDF解析結果保持 ==========")
                logger.info("現在保持PDF数 : %d", len(self.pdf_results))

                for index, item in enumerate(self.pdf_results, start=1):

                    logger.info(
                        "[%d] PDF : %s",
                        index,
                        os.path.basename(item["pdf_path"])
                    )

                    logger.info(
                        "[%d] 品名 : %s",
                        index,
                        item["data"].get("品名", "")
                    )

                    logger.info(
                        "[%d] 金額 : %s",
                        index,
                        item["data"].get("金額", "")
                    )


                if result:


                    subject = result.get(
                        "品名",
                        ""
                    )

                    amount = result.get(
                        "金額",
                        ""
                    )


                    logger.info(
                        "品名 : %s",
                        subject
                    )

                    logger.info(
                        "金額 : %s",
                        amount
                    )


            self.set_status(
                "PDF解析完了",
                100
            )


            QMessageBox.information(
                self,
                "完了",
                f"{len(self.pdf_results)}件解析しました"
            )



        except Exception:


            logger.exception(
                "TG PDF解析失敗"
            )

    """
    PDF名から対応するExcelを検索する

    例

    仕様書_QY6247.pdf
        ↓

    QY6247 をキーとして

    QY6247-9452.xlsm

    を探す
    """
    def find_excel_for_pdf(
        self,
        pdf_path: str,
        excel_files: list[str]
    ) -> str | None:
        """
        PDFに対応するExcelを検索

        例
        仕様書_QY6247.pdf
            ↓
        QY6247-9452.xlsm
        """

        pdf_name = os.path.basename(pdf_path)

        # QY6247 を取得
        key = pdf_name.replace("仕様書_", "")
        key = os.path.splitext(key)[0]

        logger.info(
            "検索キー : %s",
            key
        )

        for excel in excel_files:

            excel_name = os.path.basename(excel)

            if key in excel_name:

                logger.info(
                    "対応Excel : %s",
                    excel_name
                )

                return excel

        logger.warning(
            "対応Excelなし : %s",
            pdf_name
        )

        return None
        
    """
    台帳へ記入する

    処理

    OCR結果保持リスト
        ↓

    豊田合成シートへ追記
        ↓

    採番した見積番号を
    pdf_resultsへ保持する
    """
    def write_ledger(self):

        if not self.pdf_results:

            QMessageBox.warning(
                self,
                "確認",
                "先にPDF解析してください。"
            )

            return

        #
        # 仮ローカルpath
        #
        ledger_path = r"D:\AIM\202606\20260713見積書自動化\台帳\2026年_見積・請求・注文書発行管理台帳(その他)_ES部.xlsx"

        writer = TgLedgerWriter()

        writer.write(
            ledger_path,
            self.pdf_results
        )

        QMessageBox.information(
            self,
            "完了",
            "台帳記入が完了しました。"
        )

    """
    Excel・PDF一括出力

    処理

    保持済みOCR結果
        ↓
    入力Excelコピー
        ↓
    OCR結果入力
        ↓
    見積番号入力
        ↓
    Excel保存
        ↓
    PDF変換
    """
    def export_excel_pdf(self):

        # PDF解析済み確認
        if not hasattr(self, "pdf_results") or not self.pdf_results:

            QMessageBox.warning(
                self,
                "確認",
                "先にPDF解析を実行してください。"
            )
            return

        # 出力フォルダ
        output_dir = self.output_dir_edit.text().strip()

        if not output_dir:

            QMessageBox.warning(
                self,
                "確認",
                "出力フォルダを選択してください。"
            )
            return

        os.makedirs(
            output_dir,
            exist_ok=True
        )

        writer = TgExcelWriter()

        self.progress.setMaximum(
            len(self.pdf_results)
        )

        self.progress.setValue(0)

        self.output_button.setEnabled(False)

        self.add_log("")
        self.add_log("========== Excel・PDF出力開始 ==========")
        self.add_log(f"出力先 : {output_dir}")

        success = 0

        try:
            
            # PDFごとにExcel・PDF出力
            for index, result in enumerate(
                self.pdf_results,
                start=1
            ):

                pdf_path = result["pdf_path"]

                input_excel = result["excel_path"]

                data = result["data"]

                estimate_no = result["estimate_no"]

                filename = os.path.splitext(
                    os.path.basename(pdf_path)
                )[0]

                output_excel = os.path.join(
                    output_dir,
                    filename + ".xlsm"
                )

                output_pdf = os.path.join(
                    output_dir,
                    filename + ".pdf"
                )

                self.set_status(
                    f"出力中 ({index}/{len(self.pdf_results)})",
                    index
                )

                self.add_log("")
                self.add_log(
                    f"処理 : {filename}"
                )

                self.add_log(
                    f"元Excel : {os.path.basename(input_excel)}"
                )

                self.add_log(
                    f"見積番号 : {estimate_no}"
                )

                # ExcelへOCR結果を書き込む
                writer.write(
                    input_excel_path=input_excel,
                    output_excel_path=output_excel,
                    data=data,
                    estimate_no=estimate_no
                )

                # ExcelをPDFへ変換
                writer.export_pdf(
                    output_excel,
                    output_pdf
                )

                self.add_log(
                    f"Excel : {os.path.basename(output_excel)}"
                )

                self.add_log(
                    f"PDF   : {os.path.basename(output_pdf)}"
                )

                success += 1

                self.progress.setValue(index)

            self.set_status(
                "Excel・PDF出力完了",
                len(self.pdf_results)
            )

            self.add_log("")
            self.add_log("========== 出力終了 ==========")
            self.add_log(
                f"成功 : {success} 件"
            )

            QMessageBox.information(
                self,
                "完了",
                f"{success}件のExcel・PDFを出力しました。"
            )

        except Exception as e:

            logger.exception(
                "Excel・PDF出力失敗"
            )

            self.add_log(
                f"ERROR : {e}"
            )

            QMessageBox.critical(
                self,
                "エラー",
                str(e)
            )

        # 出力ボタンを再度有効化
        finally:

            self.output_button.setEnabled(True)