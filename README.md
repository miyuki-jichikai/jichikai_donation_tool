# 寄付依頼書作成ツール
## Jichikai Donation Letter Generator

三幸町自治会 企業寄付依頼書 自動生成ツール

---

## フォルダ構成

```
Jichikai_donation_tool/
  donation_letter_generator.py   ← メインスクリプト
  build_exe.py                   ← EXEファイル作成スクリプト
  install_shortcut.py            ← スタートメニュー登録スクリプト
  README.md                      ← このファイル
  data/                          ← 入力ファイル置き場（GitHubにはアップしない）
    （Excelファイル）
    （PDFファイル）
    （Wordテンプレートファイル）
  output/                        ← 出力ファイル置き場（自動生成）
  tmp_ad_images/                 ← 一時画像フォルダ（自動生成）
```

---

## 必要な環境

- Windows 11
- Python 3.x
- 必要ライブラリ（以下でインストール）

```
pip install pdfplumber pillow python-docx openpyxl pdf2image pyinstaller
```

- poppler（PATH設定済みであること）
  - https://github.com/oschwartz10612/poppler-windows/releases

---

## 初回セットアップ

### 1. ファイルを配置する

`data/` フォルダに以下のファイルを置く：

| ファイル名 | 内容 |
|-----------|------|
| `250831_プログラム広告掲載リストTEST.xlsx` | 企業情報・ページ番号一覧 |
| `ad_list.pdf` | 昨年のプログラム（広告掲載）|
| `企業寄付申込書.dotx` | Wordテンプレート |

### 2. スクリプトを実行する

```
python donation_letter_generator.py
```

---

## 使い方

1. スクリプトを起動すると、Excel から処理対象企業（町内寄付・町外寄付）を読み込みます
2. 1社ずつGUI画面が開き、PDFのページが表示されます
3. マウスでドラッグして広告の範囲を選択します
4. ボタンを押して操作を選択：
   - **✅ この範囲でOK** → Wordに貼り付けて次の会社へ
   - **⏭ この会社をスキップ** → この会社を飛ばして次へ
   - **💾 中断して保存** → 途中保存して終了
5. 全件完了後、`output/寄付依頼書_完成.docx` が生成されます

### 途中再開

中断した場合、次回起動時に「続きから再開しますか？」と聞いてきます。

---

## EXE化してスタートメニューに登録する手順

### Step 1: EXEファイルを作成
```
python build_exe.py
```
→ `dist/donation_letter_generator.exe` が生成されます

### Step 2: スタートメニューに登録
```
python install_shortcut.py
```
→ スタートメニューに「**寄付依頼書作成ツール**」として登録されます

---

## Excelファイルの列構成

| 列 | 内容 |
|----|------|
| A列 | 種別（町内寄付・町外寄付・特別寄付） |
| B列 | ID |
| E列 | 会社名・店名・氏名 |
| F列 | プログラムページ番号 |
| G列 | 住所 |
| J列 | 寄付金額（年号は毎年変更） |

## 広告サイズ対応表

| 寄付金額 | 広告サイズ（幅×高さ） |
|---------|-------------------|
| 100,000円・50,000円 | 160mm × 240mm |
| 30,000円 | 150mm × 115mm |
| 20,000円 | 150mm × 75mm |
| 10,000円 | 150mm × 45mm |
| 5,000円  | 75mm × 45mm |
| 3,000円  | 75mm × 35mm |

---

## 年度更新時の注意

毎年、Excelの J列の列名（例：「25年寄付金」）が変わります。
`donation_letter_generator.py` の以下の行を更新してください：

```python
df['25年寄付金'] = pd.to_numeric(df['25年寄付金'], ...)
```

↓ 例：来年は

```python
df['26年寄付金'] = pd.to_numeric(df['26年寄付金'], ...)
```

---

## お問い合わせ

三幸町自治会
