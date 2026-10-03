from __future__ import annotations

import argparse
from pathlib import Path

from pypdf import PdfReader


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("source", type=Path)
    parser.add_argument("output", type=Path)
    parser.add_argument("--title", required=True)
    parser.add_argument("--revision", required=True)
    parser.add_argument("--official-page", required=True)
    parser.add_argument("--pdf-url", required=True)
    args = parser.parse_args()

    reader = PdfReader(args.source)
    lines = [
        "---",
        f'title: "{args.title}"',
        f'revision: "{args.revision}"',
        'publisher: "国土交通省 大臣官房 官庁営繕部"',
        'retrieved: "2026-09-14"',
        f'official_page: "{args.official_page}"',
        f'source_pdf_url: "{args.pdf_url}"',
        f'pages: {len(reader.pages)}',
        'conversion: "pypdfによるページ単位のテキスト抽出"',
        "---",
        "",
        f"# {args.title}",
        "",
        "> 国土交通省の公式PDFを、ローカル検索・AI参照用にMarkdownへ変換したものです。表組みや図版、段組みは崩れる場合があるため、判断時は原典PDFも確認してください。",
        "",
        "## 出典情報",
        "",
        f"- 版: {args.revision}",
        f"- 公式掲載ページ: {args.official_page}",
        f"- 原典PDF: {args.pdf_url}",
        f"- 取得日: 2026-09-14",
        f"- 原典ページ数: {len(reader.pages)}",
        "",
        "## 本文",
    ]

    for page_number, page in enumerate(reader.pages, start=1):
        text = (page.extract_text() or "").replace("\x00", "").strip()
        lines.extend(["", "---", "", f"## PDF {page_number}ページ", "", text or "*[テキストを抽出できませんでした]*"])

    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text("\n".join(lines) + "\n", encoding="utf-8")


if __name__ == "__main__":
    main()
