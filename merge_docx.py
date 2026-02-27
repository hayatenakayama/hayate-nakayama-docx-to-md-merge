#!/usr/bin/env python3
"""
DOCX 結合スクリプト

指定ディレクトリ内の .docx ファイルをすべて1つの Word ファイルに結合します。

使い方:
    python merge_docx.py <入力ディレクトリ> [オプション]

例:
    python merge_docx.py ./ドキュメント選別
    python merge_docx.py ./ドキュメント選別 -o merged.docx
"""

import argparse
import sys
from pathlib import Path

from docxcompose.composer import Composer
from docx import Document


def collect_docx_files(input_dir: Path) -> list[Path]:
    """
    ディレクトリ内の .docx ファイルを収集し、ファイル名順にソートして返す。
    隠しファイル（~$ で始まる一時ファイル）は除外する。
    """
    docx_files = [
        f
        for f in input_dir.glob("*.docx")
        if not f.name.startswith("~$")
    ]
    docx_files.sort(key=lambda p: p.name)
    return docx_files


def merge_docx_files(
    input_dir: Path,
    output_file: Path,
) -> None:
    """
    ディレクトリ内の全 .docx ファイルを1つの Word ファイルに結合して出力。

    Args:
        input_dir: .docx ファイルが格納されたディレクトリ
        output_file: 出力先の .docx ファイルパス
    """
    docx_files = collect_docx_files(input_dir)

    if not docx_files:
        print(f"❌ '{input_dir}' に .docx ファイルが見つかりませんでした。", file=sys.stderr)
        sys.exit(1)

    print(f"📂 入力ディレクトリ: {input_dir}")
    print(f"📄 検出ファイル数: {len(docx_files)}")
    print(f"📝 出力先: {output_file}")
    print()

    # 最初のファイルをベースにする
    print(f"  [  1/{len(docx_files)}] ベース: {docx_files[0].name}")
    base_doc = Document(str(docx_files[0]))
    composer = Composer(base_doc)

    # 2つ目以降を結合
    for i, docx_path in enumerate(docx_files[1:], 2):
        print(f"  [{i:3d}/{len(docx_files)}] 結合中: {docx_path.name}")
        try:
            doc = Document(str(docx_path))
            composer.append(doc)
        except Exception as e:
            print(f"  ❌ エラー: {docx_path.name} - {e}", file=sys.stderr)

    output_file.parent.mkdir(parents=True, exist_ok=True)
    composer.save(str(output_file))

    print()
    print(f"✅ 完了! {len(docx_files)} ファイルを結合しました → {output_file}")


def main():
    parser = argparse.ArgumentParser(
        description="DOCX ファイルを1つの Word ファイルに結合します。",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
使用例:
  python merge_docx.py ./ドキュメント選別
  python merge_docx.py ./ドキュメント選別 -o merged.docx
        """,
    )

    parser.add_argument(
        "input_dir",
        type=str,
        help="結合対象の .docx ファイルが格納されたディレクトリパス",
    )
    parser.add_argument(
        "-o", "--output",
        type=str,
        default=None,
        help="出力ファイル名 (デフォルト: <入力ディレクトリ名>_merged.docx)",
    )

    args = parser.parse_args()

    input_dir = Path(args.input_dir)
    if not input_dir.is_dir():
        print(f"❌ '{input_dir}' はディレクトリではありません。", file=sys.stderr)
        sys.exit(1)

    if args.output:
        output_file = Path(args.output)
    else:
        output_file = Path(f"{input_dir.name}_merged.docx")

    merge_docx_files(
        input_dir=input_dir,
        output_file=output_file,
    )


if __name__ == "__main__":
    main()
