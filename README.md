# 寄付依頼書作成ツール
## Jichikai Donation Letter Generator

三幸町自治会 企業寄付依頼書 自動生成ツール（現在バージョン: v2.8）

---

## フォルダ構成

```
Jichikai_donation_tool/
  donation_letter_generator.py   ← メインスクリプト
  installer.iss                  ← Inno Setup インストーラースクリプト
  version_info.txt               ← EXEのバージョン情報（PyInstaller用）
  app_icon_final.ico             ← アプリアイコン
  README.md                      ← このファイル
  data/                          ← 入力ファイル置き場（GitHubにはアップしない）
    （Excelファイル）
    （PDFファイル）
    （Wordテンプレートファイル）
  output/                        ← 出力ファイル置き場（自動生成）
  tmp_ad_images/                 ← 一時画像フォルダ（自動生成）
  progress.json                  ← 中断・再開用の進捗ファイル（自動生成）
```

---

## 一般ユーザー向け：インストール方法

開発環境を用意しなくても、以下の手順でそのまま使用できます。

1. [Releases](https://github.com/miyuki-jichikai/jichikai_donation_tool/releases) ページから最新版の `jichikai_donation_tool_setup.exe` をダウンロード
2. ダウンロードしたインストーラーを実行（管理者権限は不要です）
3. インストール後、スタートメニューまたはデスクトップのショートカットから起動

---

## 開発環境（コードを修正する場合）

- Windows 11
- Python 3.x
- 必要ライブラリ（以下でインストール）

```
pip install pdfplumber pillow python-docx openpyxl pdf2image pyinstaller
```

- poppler（PATH設定済みであること）
  - https://github.com/oschwartz10612/poppler-windows/releases

### ファイルを配置する

`data/` フォルダに以下のファイルを置く：

| ファイル名 | 内容 |
|-----------|------|
| `（年月日）_プログラム広告掲載リストTEST.xlsx` | 企業情報・ページ番号一覧 |
| `ad_list.pdf` | 昨年のプログラム（広告掲載） |
| `企業寄付申込書.dotx` | Wordテンプレート |

### スクリプトを実行する

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

## EXE化・インストーラー作成手順（開発者向け）

### Step 1: バージョン番号を更新

以下3箇所のバージョン番号を揃えて更新する：

- `donation_letter_generator.py` の `APP_VERSION`
- `version_info.txt` の `filevers` / `prodvers` / `FileVersion` / `ProductVersion`
- `installer.iss` の `MyAppVersion`

### Step 2: EXEファイルを作成（PyInstaller）

```powershell
python -m PyInstaller `
  --onefile `
  --windowed `
  --icon="app_icon_final.ico" `
  --name="jichikai_donation_tool" `
  --version-file="version_info.txt" `
  --noupx `
  donation_letter_generator.py
```

→ `dist/jichikai_donation_tool.exe` が生成されます

### Step 3: インストーラーを作成（Inno Setup）

`installer.iss` をInno Setup Compilerで開き、「Build」→「Compile」を実行する。

→ `Output/jichikai_donation_tool_setup.exe` が生成されます

### Step 4: GitHubへ公開

```powershell
git add .
git commit -m "vX.X: 変更内容"
git push
```

GitHubの [Releases](https://github.com/miyuki-jichikai/jichikai_donation_tool/releases) ページから新しいタグ・Releaseを作成し、`jichikai_donation_tool_setup.exe` をAssetとして添付して公開する。

---

## Excelファイルの列構成（v2.8時点）

| 列 | 内容 |
|----|------|
| A列 | ID |
| C列 | 会社名・店名・氏名 |
| G列 | 寄付金額（年号は毎年変更） |
| X列 | 敬称 |
| AB列 | プログラムページ番号 |

※列の位置は年度によって変更される場合があります。実際にスクリプトを実行し、コンソール（VSCodeのターミナル）に表示される列名を確認してください。

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

毎年、Excelの寄付金額列の列名（例：「25年寄付金」）が変わります。
`donation_letter_generator.py` は列名ではなく**列の位置（G列）**で寄付金額を取得する仕様のため、列の位置が変わらない限り、コードの変更は不要です。列の位置自体が変わった場合は、列番号（0始まりのインデックス）を修正してください。

---

## お問い合わせ

三幸町自治会