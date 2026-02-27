# DOCX / PDF to Markdown

指定ディレクトリ内の Word (.docx) または PDF (.pdf) ファイルを変換・結合するスクリプト集です。

- **DOCX → Markdown（個別出力）**: 各 .docx を個別の .md ファイルとして新しいフォルダに出力
- **DOCX → Markdown（結合出力）**: 変換後、1つの .md ファイルに結合して出力
- **DOCX → DOCX**: Word ファイルをそのまま1つの .docx に結合
- **PDF → Markdown**: 変換後、個別の .md ファイルとして出力
- **GAS**: Google ドキュメントから不要タブを除去して別フォルダへ出力

## セットアップ

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## 使い方

### DOCX → Markdown（個別出力）

```bash
# 基本（元フォルダ名_md に個別の .md ファイルを出力、画像は Base64 埋め込み）
python convert_docx.py ./ドキュメント選別

# 出力先ディレクトリを指定
python convert_docx.py ./ドキュメント選別 -o ./output_md

# 画像をファイルとして抽出
python convert_docx.py ./ドキュメント選別 -o ./output_md --images-dir ./images
```

### DOCX → Markdown（結合出力）

```bash
# 基本（画像は Base64 埋め込み）
python convert_and_merge.py ./ドキュメント選別

# 出力ファイル名を指定
python convert_and_merge.py ./ドキュメント選別 -o merged.md

# 画像をファイルとして抽出
python convert_and_merge.py ./ドキュメント選別 -o merged.md --images-dir ./images
```

### DOCX 結合（Word → Word）

```bash
# 基本
python merge_docx.py ./ドキュメント選別

# 出力ファイル名を指定
python merge_docx.py ./ドキュメント選別 -o merged.docx
```

### PDF → Markdown（個別出力）

```bash
# 基本（各PDFが個別の .md ファイルとして出力される）
python convert_pdf_and_merge.py ./PDFフォルダ

# 出力先ディレクトリを指定
python convert_pdf_and_merge.py ./PDFフォルダ -o ./output

# 画像をファイルとして抽出
python convert_pdf_and_merge.py ./PDFフォルダ -o ./output --images-dir ./images
```

### Google ドキュメント 不要タブ除去（GAS）

`週次会議_evaluate_optimizer.js` は Google Apps Script 用のスクリプトです。
指定フォルダ内の Google ドキュメントから不要なタブを除去し、別フォルダへ出力します。

1. [Google Apps Script](https://script.google.com/) で新しいプロジェクトを作成
2. `週次会議_evaluate_optimizer.js` の内容を貼り付け
3. `CONFIG` の `INPUT_FOLDER_ID` と `OUTPUT_FOLDER_ID` を対象の Google Drive フォルダ ID に変更
4. `Optimize_Weekly_MTG` 関数を実行

## オプション

### DOCX 個別出力 (`convert_docx.py`)

| オプション | 説明 | デフォルト |
|---|---|---|
| `input_dir` | 変換対象ディレクトリ（必須） | — |
| `-o`, `--output` | 出力先ディレクトリ | `<ディレクトリ名>_md` |
| `--images-dir` | 画像の保存先ディレクトリ | なし（Base64 埋め込み） |

### DOCX 結合出力 (`convert_and_merge.py`)

| オプション | 説明 | デフォルト |
|---|---|---|
| `input_dir` | 変換対象ディレクトリ（必須） | — |
| `-o`, `--output` | 出力ファイル名 | `<ディレクトリ名>_merged.md` |
| `--images-dir` | 画像の保存先ディレクトリ | なし（Base64 埋め込み） |
| `--separator` | ドキュメント間の区切り | `---` |

### DOCX 結合 (`merge_docx.py`)

| オプション | 説明 | デフォルト |
|---|---|---|
| `input_dir` | 結合対象ディレクトリ（必須） | — |
| `-o`, `--output` | 出力ファイル名 | `<ディレクトリ名>_merged.docx` |

### PDF (`convert_pdf_and_merge.py`)

| オプション | 説明 | デフォルト |
|---|---|---|
| `input_dir` | 変換対象ディレクトリ（必須） | — |
| `-o`, `--output` | 出力先ディレクトリ | `<ディレクトリ名>_pdf_md` |
| `--images-dir` | 画像の保存先ディレクトリ | なし（画像なし） |

## 出力形式

### DOCX → Markdown（個別出力）
- 各 .docx が `<ファイル名>.md` として個別に出力されます
- 出力先フォルダはデフォルトで `<入力ディレクトリ名>_md`
- ファイルはファイル名順（昇順）にソートされます
- Word の一時ファイル（`~$` で始まるファイル）は自動的に除外されます

### DOCX → Markdown（結合出力）
- 各ファイルがファイル名を `# 見出し` として1つの .md ファイルに結合されます
- ファイルはファイル名順（昇順）にソートされます
- Word の一時ファイル（`~$` で始まるファイル）は自動的に除外されます

### DOCX 結合
- 複数の .docx ファイルを1つの .docx ファイルに結合します
- ファイルはファイル名順（昇順）にソートされます
- Word の一時ファイル（`~$` で始まるファイル）は自動的に除外されます

### PDF
- 各PDFが `<ファイル名>.md` として個別に出力されます
- ファイルはファイル名順（昇順）に処理されます

## 依存ライブラリ

| ライブラリ | 用途 |
|---|---|
| [mammoth](https://pypi.org/project/mammoth/) | DOCX → HTML 変換 |
| [markdownify](https://pypi.org/project/markdownify/) | HTML → Markdown 変換 |
| [pymupdf4llm](https://pypi.org/project/pymupdf4llm/) | PDF → Markdown 変換 |
| [docxcompose](https://pypi.org/project/docxcompose/) | DOCX 結合 |
