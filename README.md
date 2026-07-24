# 見積書作成ツール

業務計画書・見積依頼書などから必要な情報を抽出し、
各社指定フォーマットの見積書および管理台帳を自動作成する
Windows向けデスクトップアプリです。

---

## 主な機能

### 特調TB

- 業務計画書(PDF)から情報抽出
- 管理台帳検索
- TB見積書作成
- TB管理台帳更新

### 特調TB以外

- 業務計画書(PDF)から情報抽出
- 管理台帳検索
- TB見積書作成
- TB管理台帳更新

### TG

- TG PDFから情報抽出
- AI(OpenAI Vision)による情報補完
- TG見積書作成
- TG管理台帳更新

### MSR

- MSR入力シート読込
- MSR見積書作成
- MSR管理台帳更新

### 共通機能

- 一括実行
- 抽出結果プレビュー
- 処理進捗表示
- エラー表示
- OneDrive対応
- exe配布対応

---

## システム構成

```
estimate-app/
├── config/
├── doc/
├── models/
├── resources/
│   ├── msr/
│   └── tb/
├── services/
│   ├── ai/
│   ├── cloud/
│   ├── excel/
│   ├── readers/
│   └── writers/
├── ui/
├── workers/
├── main.py
├── build.ps1
└── requirements.txt
```

---

## 使用技術

- Python 3.14
- PySide6
- openpyxl
- xlwings
- pdfplumber
- Microsoft Graph API
- OpenAI Vision API
- PyInstaller

---

## 配布ファイル

```
EstimateTool/
├── EstimateTool.exe
├── .env
└── resources/
```

---

## 実行方法

1. `EstimateTool.exe` を起動します。
2. 処理したいタブを選択します。
3. 必要なファイルを選択します。
4. **一括実行** をクリックします。
5. 見積書・管理台帳が自動作成されます。

---

## 操作マニュアル

詳細な操作方法は以下を参照してください。

```
doc/HowTo/HowTo.md
```

---

## ライセンス

社内利用を目的としたツールです。