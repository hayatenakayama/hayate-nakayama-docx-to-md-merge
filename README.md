# DOCX to Markdown Merge

指定ディレクトリ内の Word (.docx) ファイルを Markdown に変換し、1つのファイルに結合するスクリプトです。

## セットアップ

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## 使い方

```bash
# 基本（画像は Base64 埋め込み）
python convert_and_merge.py ./ドキュメント選別

# 出力ファイル名を指定
python convert_and_merge.py ./ドキュメント選別 -o merged.md

# 画像をファイルとして抽出
python convert_and_merge.py ./ドキュメント選別 -o merged.md --images-dir ./images

# ドキュメント間の区切りを変更
python convert_and_merge.py ./ドキュメント選別 --separator "==="
```

## オプション

| オプション | 説明 | デフォルト |
|---|---|---|
| `input_dir` | 変換対象の `.docx` ディレクトリ（必須） | — |
| `-o`, `--output` | 出力ファイル名 | `<ディレクトリ名>_merged.md` |
| `--images-dir` | 画像の保存先ディレクトリ | なし（Base64 埋め込み） |
| `--separator` | ドキュメント間の区切り | `---` |

## 出力形式

- 各 `.docx` ファイルがファイル名を `# 見出し` として結合されます
- ファイルはファイル名順（昇順）にソートされます
- Word の一時ファイル（`~$` で始まるファイル）は自動的に除外されます

## 依存ライブラリ

- [mammoth](https://pypi.org/project/mammoth/) — DOCX → HTML 変換
- [markdownify](https://pypi.org/project/markdownify/) — HTML → Markdown 変換
