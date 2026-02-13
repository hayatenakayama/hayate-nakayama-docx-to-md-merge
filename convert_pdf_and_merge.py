#!/usr/bin/env python3
"""
PDF → Markdown 変換スクリプト

指定ディレクトリ内の .pdf ファイルをすべて Markdown に変換し、
個別の .md ファイルとして出力します。

使い方:
    python convert_pdf_and_merge.py <入力ディレクトリ> [オプション]

例:
    python convert_pdf_and_merge.py ./PDFフォルダ
    python convert_pdf_and_merge.py ./PDFフォルダ -o ./output --images-dir images
"""

import argparse
import os
import sys
from pathlib import Path
from datetime import datetime

import pymupdf4llm
import pymupdf


def convert_pdf_to_markdown(
    pdf_path: Path,
    images_dir: Path | None = None,
) -> str:
    """
    単一の .pdf ファイルを Markdown 文字列に変換する。

    Args:
        pdf_path: 変換対象の .pdf ファイルパス
        images_dir: 画像の出力先ディレクトリ（None の場合、画像は出力しない）

    Returns:
        変換後の Markdown 文字列
    """
    if images_dir:
        images_dir.mkdir(parents=True, exist_ok=True)
        markdown = pymupdf4llm.to_markdown(
            str(pdf_path),
            write_images=True,
            image_path=str(images_dir),
        )
    else:
        markdown = pymupdf4llm.to_markdown(str(pdf_path))

    # 余分な空行を整理
    lines = markdown.split("\n")
    cleaned_lines = []
    prev_empty = False
    for line in lines:
        is_empty = line.strip() == ""
        if is_empty and prev_empty:
            continue
        cleaned_lines.append(line)
        prev_empty = is_empty

    return "\n".join(cleaned_lines).strip()


def collect_pdf_files(input_dir: Path) -> list[Path]:
    """
    ディレクトリ内の .pdf ファイルを収集し、ファイル名順にソートして返す。
    """
    pdf_files = list(input_dir.glob("*.pdf"))
    pdf_files.sort(key=lambda p: p.name)
    return pdf_files


def convert_all_pdfs(
    input_dir: Path,
    output_dir: Path,
    images_dir: Path | None = None,
) -> None:
    """
    ディレクトリ内の全 .pdf ファイルを変換し、個別の Markdown ファイルとして出力。

    Args:
        input_dir: .pdf ファイルが格納されたディレクトリ
        output_dir: 出力先ディレクトリ
        images_dir: 画像の出力先ディレクトリ
    """
    pdf_files = collect_pdf_files(input_dir)

    if not pdf_files:
        print(f"❌ '{input_dir}' に .pdf ファイルが見つかりませんでした。", file=sys.stderr)
        sys.exit(1)

    output_dir.mkdir(parents=True, exist_ok=True)

    print(f"📂 入力ディレクトリ: {input_dir}")
    print(f"📄 検出ファイル数: {len(pdf_files)}")
    print(f"� 出力先ディレクトリ: {output_dir}")
    print()

    success_count = 0

    for i, pdf_path in enumerate(pdf_files, 1):
        print(f"  [{i:3d}/{len(pdf_files)}] 変換中: {pdf_path.name}")
        try:
            markdown_content = convert_pdf_to_markdown(pdf_path, images_dir)

            output_file = output_dir / f"{pdf_path.stem}.md"
            output_file.write_text(markdown_content, encoding="utf-8")
            success_count += 1
        except Exception as e:
            print(f"  ❌ エラー: {pdf_path.name} - {e}", file=sys.stderr)

    print()
    print(f"✅ 完了! {success_count}/{len(pdf_files)} ファイルを変換しました → {output_dir}")


def main():
    parser = argparse.ArgumentParser(
        description="PDF ファイルを Markdown に変換し、個別のファイルとして出力します。",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
使用例:
  python convert_pdf_and_merge.py ./PDFフォルダ
  python convert_pdf_and_merge.py ./PDFフォルダ -o ./output
  python convert_pdf_and_merge.py ./PDFフォルダ -o ./output --images-dir ./images
        """,
    )

    parser.add_argument(
        "input_dir",
        type=str,
        help="変換対象の .pdf ファイルが格納されたディレクトリパス",
    )
    parser.add_argument(
        "-o", "--output",
        type=str,
        default=None,
        help="出力先ディレクトリ (デフォルト: <入力ディレクトリ名>_pdf_md)",
    )
    parser.add_argument(
        "--images-dir",
        type=str,
        default=None,
        help="画像の保存先ディレクトリ (未指定の場合、画像は出力しない)",
    )

    args = parser.parse_args()

    input_dir = Path(args.input_dir)
    if not input_dir.is_dir():
        print(f"❌ '{input_dir}' はディレクトリではありません。", file=sys.stderr)
        sys.exit(1)

    if args.output:
        output_dir = Path(args.output)
    else:
        output_dir = Path(f"{input_dir.name}_pdf_md")

    images_dir = Path(args.images_dir) if args.images_dir else None

    convert_all_pdfs(
        input_dir=input_dir,
        output_dir=output_dir,
        images_dir=images_dir,
    )


if __name__ == "__main__":
    main()
