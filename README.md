# DOCX / PDF to Markdown

指定ディレクトリ内の Word (.docx) または PDF (.pdf) ファイルを Markdown に変換するスクリプトです。

- **DOCX**: 変換後、1つのファイルに結合して出力
- **PDF**: 変換後、個別の .md ファイルとして出力

## セットアップ

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## 使い方

### DOCX → Markdown（結合出力）

```bash
# 基本（画像は Base64 埋め込み）
python convert_and_merge.py ./ドキュメント選別

# 出力ファイル名を指定
python convert_and_merge.py ./ドキュメント選別 -o merged.md

# 画像をファイルとして抽出
python convert_and_merge.py ./ドキュメント選別 -o merged.md --images-dir ./images
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

## オプション

### DOCX (`convert_and_merge.py`)

| オプション | 説明 | デフォルト |
|---|---|---|
| `input_dir` | 変換対象ディレクトリ（必須） | — |
| `-o`, `--output` | 出力ファイル名 | `<ディレクトリ名>_merged.md` |
| `--images-dir` | 画像の保存先ディレクトリ | なし（Base64 埋め込み） |
| `--separator` | ドキュメント間の区切り | `---` |

### PDF (`convert_pdf_and_merge.py`)

| オプション | 説明 | デフォルト |
|---|---|---|
| `input_dir` | 変換対象ディレクトリ（必須） | — |
| `-o`, `--output` | 出力先ディレクトリ | `<ディレクトリ名>_pdf_md` |
| `--images-dir` | 画像の保存先ディレクトリ | なし（画像なし） |

## 出力形式

### DOCX
- 各ファイルがファイル名を `# 見出し` として1つのファイルに結合されます
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
