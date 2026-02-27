#!/usr/bin/env python3
"""
DOCX → Markdown 個別変換スクリプト

指定ディレクトリ内の .docx ファイルをすべて Markdown に変換し、
個別の .md ファイルとして新しいフォルダに出力します。

使い方:
    python convert_docx.py <入力ディレクトリ> [オプション]

例:
    python convert_docx.py ./ドキュメント選別
    python convert_docx.py ./ドキュメント選別 -o ./output_md
    python convert_docx.py ./ドキュメント選別 --images-dir images
"""

import argparse
import sys
import base64
from pathlib import Path
from datetime import datetime

import mammoth
from markdownify import markdownify as md


def convert_docx_to_markdown(
    docx_path: Path,
    images_dir: Path | None = None,
) -> str:
    """
    単一の .docx ファイルを Markdown 文字列に変換する。

    Args:
        docx_path: 変換対象の .docx ファイルパス
        images_dir: 画像の出力先ディレクトリ（None の場合、画像は Base64 で埋め込み）

    Returns:
        変換後の Markdown 文字列
    """
    image_counter = {"count": 0}
    stem = docx_path.stem

    def convert_image(image):
        """mammoth の画像ハンドラ: 画像をファイルに保存し Markdown リンクを返す"""
        image_counter["count"] += 1
        ext = image.content_type.split("/")[-1]
        if ext == "jpeg":
            ext = "jpg"

        if images_dir:
            images_dir.mkdir(parents=True, exist_ok=True)
            image_filename = f"{stem}_img{image_counter['count']:03d}.{ext}"
            image_path = images_dir / image_filename

            with image.open() as image_bytes:
                image_path.write_bytes(image_bytes.read())

            return {"src": str(image_path)}
        else:
            # 画像ディレクトリ未指定の場合は Base64 埋め込み
            with image.open() as image_bytes:
                data = base64.b64encode(image_bytes.read()).decode("utf-8")
            return {"src": f"data:{image.content_type};base64,{data}"}

    # DOCX → HTML 変換
    with open(docx_path, "rb") as f:
        result = mammoth.convert_to_html(
            f,
            convert_image=mammoth.images.img_element(convert_image),
        )

    html = result.value

    # 変換時の警告を表示
    if result.messages:
        for msg in result.messages:
            print(f"  ⚠ {msg}", file=sys.stderr)

    # HTML → Markdown 変換
    markdown = md(
        html,
        heading_style="atx",
        bullets="-",
        strip=["script", "style"],
    )

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


def convert_all_docx(
    input_dir: Path,
    output_dir: Path,
    images_dir: Path | None = None,
) -> None:
    """
    ディレクトリ内の全 .docx ファイルを変換し、個別の Markdown ファイルとして出力。

    Args:
        input_dir: .docx ファイルが格納されたディレクトリ
        output_dir: 出力先ディレクトリ
        images_dir: 画像の出力先ディレクトリ
    """
    docx_files = collect_docx_files(input_dir)

    if not docx_files:
        print(f"❌ '{input_dir}' に .docx ファイルが見つかりませんでした。", file=sys.stderr)
        sys.exit(1)

    output_dir.mkdir(parents=True, exist_ok=True)

    print(f"📂 入力ディレクトリ: {input_dir}")
    print(f"📄 検出ファイル数: {len(docx_files)}")
    print(f"📁 出力先ディレクトリ: {output_dir}")
    print()

    success_count = 0

    for i, docx_path in enumerate(docx_files, 1):
        print(f"  [{i:3d}/{len(docx_files)}] 変換中: {docx_path.name}")
        try:
            markdown_content = convert_docx_to_markdown(docx_path, images_dir)

            output_file = output_dir / f"{docx_path.stem}.md"
            output_file.write_text(markdown_content, encoding="utf-8")
            success_count += 1
        except Exception as e:
            print(f"  ❌ エラー: {docx_path.name} - {e}", file=sys.stderr)

    print()
    print(f"✅ 完了! {success_count}/{len(docx_files)} ファイルを変換しました → {output_dir}")


def main():
    parser = argparse.ArgumentParser(
        description="DOCX ファイルを Markdown に変換し、個別のファイルとして出力します。",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
使用例:
  python convert_docx.py ./ドキュメント選別
  python convert_docx.py ./ドキュメント選別 -o ./output_md
  python convert_docx.py ./ドキュメント選別 -o ./output_md --images-dir ./images
        """,
    )

    parser.add_argument(
        "input_dir",
        type=str,
        help="変換対象の .docx ファイルが格納されたディレクトリパス",
    )
    parser.add_argument(
        "-o", "--output",
        type=str,
        default=None,
        help="出力先ディレクトリ (デフォルト: <入力ディレクトリ名>_md)",
    )
    parser.add_argument(
        "--images-dir",
        type=str,
        default=None,
        help="画像の保存先ディレクトリ (未指定の場合、画像は Base64 で埋め込み)",
    )

    args = parser.parse_args()

    input_dir = Path(args.input_dir)
    if not input_dir.is_dir():
        print(f"❌ '{input_dir}' はディレクトリではありません。", file=sys.stderr)
        sys.exit(1)

    if args.output:
        output_dir = Path(args.output)
    else:
        output_dir = Path(f"{input_dir.name}_md")

    images_dir = Path(args.images_dir) if args.images_dir else None

    convert_all_docx(
        input_dir=input_dir,
        output_dir=output_dir,
        images_dir=images_dir,
    )


if __name__ == "__main__":
    main()
