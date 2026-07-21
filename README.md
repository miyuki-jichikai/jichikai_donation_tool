# 寄付依頼書作成ツール
## Jichikai Donation Letter Generator

三幸町自治会 企業寄付依頼書 自動生成ツール（現在バージョン: v2.9）

---

## フォルダ構成

```
Jichikai_donation_tool/
  donation_letter_generator.py   ← メインスクリプト
  installer.iss                  ← Inno Setup インストーラースクリプト
  version_info.txt               ← EXEのバージョン情報（PyInstaller用）
  app_icon_final.ico             ← アプリアイコン
  README.md                      ← このファイル
  poppler_bin/                   ← poppler本体（PDF→画像変換に必要。GitHubにはアップしない／
                                     配布時にEXEと同じフォルダに同梱する）
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

1. [Releases](https://github.com/miyuki-jichikai/jichikai_donation_tool/releases) ページから最新版の `jichikai_donation_tool_setup.zip` をダウンロードし、右クリック→「すべて展開」で解凍する
2. 解凍してできた `jichikai_donation_tool_setup.exe` を実行（管理者権限は不要です）
3. インストール後、スタートメニューまたはデスクトップのショートカットから起動

### ⚠ インストール時に警告が出た場合

このインストーラーは個人・自治会内での配布用のため、Microsoftへの発行元登録（コード署名）を行っていません。
そのため、インストーラーを実行すると Windows から次のような警告が表示されることがあります。

- ブラウザのダウンロードバーに「このファイルは危険な場合があります」「保持しますか？」等の表示が出た場合
  → **「保持」または「詳細情報」→「実行」**を選んでください
- 実行時に「WindowsによってPCが保護されました」という青い画面（SmartScreen）が出た場合
  → **「詳細情報」をクリック→表示された「実行」ボタン**を押してください

いずれも、既知のソフトであれば安全に進めて問題ありません。表示が不安な場合は、開発者（三幸町自治会）にご確認ください。

---

## 開発環境（コードを修正する場合）

- Windows 11
- Python 3.x
- 必要ライブラリ（以下でインストール）

```
pip install pdfplumber pillow python-docx openpyxl pdf2image pyinstaller
```

- poppler
  - https://github.com/oschwartz10612/poppler-windows/releases から最新版をダウンロード・解凍
  - 開発時：PATHを通しておけば `poppler_bin` フォルダが無くてもそのまま動作します
    （`POPPLER_PATH` はシステムPATHへ自動フォールバックします）
  - **配布用EXEを作る場合**：解凍した中の `Library\bin` フォルダを、このリポジトリ直下に
    `poppler_bin` という名前でコピーしてください（PyInstaller・Inno Setupが自動的に同梱します）。
    これにより、配布先のPCにpopplerが入っていなくてもアプリが動作します。

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

（PyInstallerのコマンド自体はv2.8までと同じです。popplerはEXEの中に
　埋め込むのではなく、次のStep 3でEXEと同じフォルダに別途配置します）

### Step 3: インストーラーを作成（Inno Setup）

事前に、リポジトリ直下に `poppler_bin` フォルダ（poppler-windows releasesの
`Library\bin` の中身）を配置しておく（開発環境セットアップの節を参照）。

`installer.iss` をInno Setup Compilerで開き、「Build」→「Compile」を実行する。
（`poppler_bin` フォルダがEXEと同じ場所に自動的に同梱されます）

→ `Output/jichikai_donation_tool_setup.exe` が生成されます

### Step 4: GitHubへ公開

公開する際は、**①ソースコードの更新**と**②インストーラー（実行ファイル）の公開**の、
性質の異なる2つの作業を行います。①はコードの保管、②はダウンロード用ファイルの公開で、
保存される場所がまったく別なので、両方とも行う必要があります。

#### ① ソースコード一式をリポジトリへpush

`.gitignore` で `poppler_bin/`・`dist/`・`Output/`・`data/` を除外しているため、
`git add .` を実行しても以下のようなソースファイルだけが対象になります
（実行ファイルや個人データは含まれません）：

- `donation_letter_generator.py`
- `installer.iss`
- `version_info.txt`
- `README.md`
- `app_icon_final.ico`

```powershell
git add .
git commit -m "v2.9: poppler同梱による配布先PC依存の解消"
git push
```

#### ② インストーラーをzip化してReleasesページで公開

**準備：exeをzip化**

1. エクスプローラーで `Output\jichikai_donation_tool_setup.exe` を右クリック
2. 「送る」→「圧縮(zip形式)フォルダー」を選択
3. 同じ場所に `jichikai_donation_tool_setup.zip` が作成される
   （zipにする理由：ブラウザがexeを未確認の実行ファイルとして警告・ブロックし、
   ダウンロード中にファイル名が変わってしまう現象を防ぐため）

**GitHub側の操作**

1. ブラウザで [Releases](https://github.com/miyuki-jichikai/jichikai_donation_tool/releases) ページを開く
2. 右上あたりにある「**Draft a new release**」（日本語UIの場合は「新しいリリースを下書き」）ボタンをクリック
3. 「**Choose a tag**」の入力欄に新しいバージョン番号、例えば `v2.9` と入力し、
   下に出てくる「**Create new tag: v2.9 on publish**」をクリックして選択
   （① で `git push` 済みの最新コードに対して、このタグが自動的に付けられます）
4. 「**Release title**」に分かりやすいタイトルを入力
   （例：`三幸町自治会 寄付依頼書生成ツール v2.9`）
5. 本文欄（Describe this release）に変更点を書く
   （例：`poppler同梱により、他のPCにインストールしても起動できない不具合を修正`）
6. 画面下の方にある点線の四角い枠「**Attach binaries by dropping them here or selecting them**」に、
   先ほど作った `jichikai_donation_tool_setup.zip` をドラッグ＆ドロップ
   （または枠内をクリックしてファイル選択ダイアログから選ぶ）
7. アップロードが完了する（zipのアイコンとファイル名が表示される）のを確認
8. 一番下の緑色の「**Publish release**」ボタンをクリックして公開する

公開後、Releasesページの一覧に新しいバージョンが表示され、そこから誰でも
`jichikai_donation_tool_setup.zip` をダウンロードできるようになります。

①と②は別々の場所に保存されます（①はリポジトリのコード、②はReleasesページのダウンロード用ファイル）。
両方行って初めて、次にツールを使う人が最新のソースを見られ、かつ最新のインストーラーをダウンロードできる状態になります。

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