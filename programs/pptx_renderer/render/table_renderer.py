"""表（`TableShape`）の描画を行うモジュール．

セル位置・セルサイズ・セル結合（gridSpan/rowSpan）・背景・枠線・テキストを
再現する．
"""

from __future__ import annotations

from reportlab.pdfgen import canvas as canvas_module

from pptx_renderer.fonts.registry import FontRegistry
from pptx_renderer.ir import TableShape
from pptx_renderer.layout.text_layout import layout_text_body
from pptx_renderer.math.latex_render import LatexMathRenderer
from pptx_renderer.render.overlay import OverlayItem
from pptx_renderer.render.text_renderer import draw_text_body
from pptx_renderer.units import CoordinateTransformer, RectEMU
from pptx_renderer.warnings_log import WarningLog


def _required_cell_height_emu(
    text_body, col_width_emu: float, font_registry: FontRegistry, math_renderer: LatexMathRenderer, warning_log: WarningLog, slide_index
) -> float:
    if text_body is None:
        return 0.0
    inset_left_pt = text_body.inset_left_emu / 12700.0
    inset_right_pt = text_body.inset_right_emu / 12700.0
    inset_top_pt = text_body.inset_top_emu / 12700.0
    inset_bottom_pt = text_body.inset_bottom_emu / 12700.0
    available_width_pt = max(1.0, col_width_emu / 12700.0 - inset_left_pt - inset_right_pt)

    lines = layout_text_body(text_body, available_width_pt, font_registry, math_renderer, warning_log, slide_index)
    content_height_pt = sum(line.space_before_pt + line.line_height_pt for line in lines)
    total_height_pt = content_height_pt + inset_top_pt + inset_bottom_pt
    return total_height_pt * 12700.0


def draw_table(
    c: canvas_module.Canvas,
    shape: TableShape,
    transformer: CoordinateTransformer,
    origin_emu: tuple[float, float],
    font_registry: FontRegistry,
    math_renderer: LatexMathRenderer,
    page_index: int,
    overlay_items: list[OverlayItem],
    warning_log: WarningLog,
    slide_index: int | None,
) -> None:
    """`TableShape`をPDFへ描画する．

    引数:
        c (canvas_module.Canvas): 描画対象のキャンバス．
        shape (TableShape): 描画対象の表．
        transformer (CoordinateTransformer): EMU→pt座標変換器．
        origin_emu (tuple[float, float]): 表の左上原点（EMU，スライド座標系）．
        font_registry (FontRegistry): フォント解決用レジストリ．
        math_renderer (LatexMathRenderer): 数式レンダラー．
        page_index (int): 出力PDFにおけるページ番号（0始まり）．
        overlay_items (list[OverlayItem]): 数式合成予約リスト（追記される）．
        warning_log (WarningLog): 警告記録先．
        slide_index (int | None): 対象スライド番号．
    戻り値:
        なし．
    """

    origin_x, origin_y = origin_emu

    col_x = [origin_x]
    for w in shape.col_widths_emu:
        col_x.append(col_x[-1] + w)

    effective_row_heights = []
    for row in shape.rows:
        required = row.height_emu
        for col_idx, cell in enumerate(row.cells):
            if cell.is_covered or cell.row_span != 1:
                continue
            col_span_width = sum(
                shape.col_widths_emu[col_idx : col_idx + cell.col_span]
            )
            required = max(
                required,
                _required_cell_height_emu(
                    cell.text_body, col_span_width, font_registry, math_renderer, warning_log, slide_index
                ),
            )
        effective_row_heights.append(required)

    row_y = [origin_y]
    for height in effective_row_heights:
        row_y.append(row_y[-1] + height)

    for row_idx, row in enumerate(shape.rows):
        for col_idx, cell in enumerate(row.cells):
            if cell.is_covered:
                continue

            x0 = col_x[col_idx]
            x1 = col_x[min(col_idx + cell.col_span, len(col_x) - 1)]
            y0 = row_y[row_idx]
            y1 = row_y[min(row_idx + cell.row_span, len(row_y) - 1)]

            cell_rect_emu = RectEMU(x=x0, y=y0, cx=x1 - x0, cy=y1 - y0)
            cell_rect_pt = transformer.rect_to_pdf(cell_rect_emu)

            if cell.fill.kind == "solid" and cell.fill.color is not None:
                c.saveState()
                c.setFillColorRGB(*cell.fill.color.to_unit_tuple())
                c.setFillAlpha(cell.fill.alpha)
                c.rect(cell_rect_pt.x, cell_rect_pt.y, cell_rect_pt.width, cell_rect_pt.height, fill=1, stroke=0)
                c.restoreState()

            _draw_border(c, cell.border_top, cell_rect_pt.x, cell_rect_pt.y + cell_rect_pt.height, cell_rect_pt.x + cell_rect_pt.width, cell_rect_pt.y + cell_rect_pt.height)
            _draw_border(c, cell.border_bottom, cell_rect_pt.x, cell_rect_pt.y, cell_rect_pt.x + cell_rect_pt.width, cell_rect_pt.y)
            _draw_border(c, cell.border_left, cell_rect_pt.x, cell_rect_pt.y, cell_rect_pt.x, cell_rect_pt.y + cell_rect_pt.height)
            _draw_border(c, cell.border_right, cell_rect_pt.x + cell_rect_pt.width, cell_rect_pt.y, cell_rect_pt.x + cell_rect_pt.width, cell_rect_pt.y + cell_rect_pt.height)

            if cell.text_body is not None:
                draw_text_body(
                    c,
                    cell.text_body,
                    cell_rect_pt,
                    font_registry,
                    math_renderer,
                    warning_log,
                    slide_index,
                    page_index,
                    overlay_items,
                )


def _draw_border(c: canvas_module.Canvas, stroke, x0: float, y0: float, x1: float, y1: float) -> None:
    if stroke.kind != "solid" or stroke.color is None:
        return
    c.saveState()
    c.setStrokeColorRGB(*stroke.color.to_unit_tuple())
    c.setStrokeAlpha(stroke.alpha)
    c.setLineWidth(stroke.width_pt)
    c.line(x0, y0, x1, y1)
    c.restoreState()
