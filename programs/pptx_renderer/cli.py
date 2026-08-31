"""コマンドラインインターフェース．

```bash
pptx-render input.pptx -o output.pdf
pptx-render input.pptx -o output.pdf --debug
pptx-render input.pptx -o output.pdf --render-pages 1,2,3
pptx-render input.pptx -o output.pdf --font-dir ./my_fonts
```
"""

from __future__ import annotations

import argparse
import sys

from pptx_renderer.fonts.registry import FontRegistry
from pptx_renderer.ir import Presentation
from pptx_renderer.parser import parse_presentation
from pptx_renderer.render.pdf_renderer import PdfRenderer
from pptx_renderer.warnings_log import WarningLog


def _parse_render_pages(value: str | None) -> list[int] | None:
    if value is None:
        return None
    return [int(v.strip()) for v in value.split(",") if v.strip()]


def _print_debug_info(presentation: Presentation, warning_log: WarningLog) -> None:
    print("=== Slide IR ===", file=sys.stderr)
    print(f"slide size: {presentation.width_emu} x {presentation.height_emu} EMU", file=sys.stderr)
    for slide in presentation.slides:
        print(f"--- slide {slide.index}: {len(slide.shapes)} shapes ---", file=sys.stderr)
        for shape in slide.shapes:
            kind = type(shape).__name__
            preset = getattr(shape, "preset", "")
            print(
                f"  [{kind}] {shape.name!r} preset={preset} "
                f"rect=({shape.rect.x:.0f},{shape.rect.y:.0f},{shape.rect.cx:.0f},{shape.rect.cy:.0f})",
                file=sys.stderr,
            )

    print("=== Warnings ===", file=sys.stderr)
    if not warning_log.entries:
        print("(warnings なし)", file=sys.stderr)
    for entry in warning_log.entries:
        print(entry.format(), file=sys.stderr)


def main(argv: list[str] | None = None) -> int:
    """CLIのエントリポイント．

    引数:
        argv (list[str] | None): コマンドライン引数．Noneの場合は`sys.argv[1:]`を使用する．
    戻り値:
        int: 終了コード（正常終了時は0）．
    """

    parser = argparse.ArgumentParser(prog="pptx-render", description="PPTXをPDFへ変換する（Linux専用レンダラー）")
    parser.add_argument("input", help="入力PPTXファイルのパス")
    parser.add_argument("-o", "--output", required=True, help="出力PDFファイルのパス")
    parser.add_argument("--debug", action="store_true", help="Slide IR・警告等のデバッグ情報を表示する")
    parser.add_argument("--render-pages", type=str, default=None, help="描画するスライド番号（例: 1,2,3）")
    parser.add_argument(
        "--workers",
        type=int,
        default=None,
        help="ページ並列描画に使用するワーカー数（既定: CPUコア数とページ数から自動決定，1で逐次実行）",
    )
    parser.add_argument(
        "--font-dir",
        action="append",
        default=None,
        help="フォント検索ディレクトリを追加する（.ttf/.otfを再帰的に検索，複数回指定可）",
    )
    parser.add_argument(
        "--no-math-cache",
        action="store_true",
        help="数式レンダリング結果のディスクキャッシュを使用しない",
    )
    args = parser.parse_args(argv)

    warning_log = WarningLog(echo_stderr=not args.debug)
    presentation = parse_presentation(args.input, warning_log)

    render_pages = _parse_render_pages(args.render_pages)
    font_registry = FontRegistry(warning_log=warning_log, extra_font_dirs=args.font_dir)
    renderer = PdfRenderer(
        warning_log=warning_log,
        font_registry=font_registry,
        use_math_disk_cache=not args.no_math_cache,
    )
    renderer.render(presentation, args.output, render_pages=render_pages, max_workers=args.workers)

    if args.debug:
        _print_debug_info(presentation, warning_log)
    elif warning_log.entries:
        print(f"警告: {len(warning_log.entries)} 件（詳細は標準エラー出力を参照）", file=sys.stderr)

    print(f"完了: {args.output}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
