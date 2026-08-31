"""レイアウト済みテキスト（`LineBox`）をPDFへ描画するモジュール．

テキストの描画そのものはReportLabのキャンバスAPIを用いて行う．
数式（`MathRun`由来のセグメント）は，この時点ではPDF上に直接描画せず，
配置位置・サイズの情報のみを`overlay_collector`へ登録する．実際の合成は
`pptx_renderer.render.overlay`が担当する（数式は独立したPDFとして生成した
ものをページへ合成するため）．
"""

from __future__ import annotations

from reportlab.pdfgen import canvas as canvas_module

from pptx_renderer.ir import HAlign, TextBody, VAlign
from pptx_renderer.layout.text_layout import LineBox, layout_text_body
from pptx_renderer.math.latex_render import LatexMathRenderer
from pptx_renderer.fonts.registry import FontRegistry
from pptx_renderer.render.overlay import OverlayItem
from pptx_renderer.units import RectPt
from pptx_renderer.warnings_log import WarningLog

_ALIGN_FACTOR = {HAlign.LEFT: 0.0, HAlign.CENTER: 0.5, HAlign.RIGHT: 1.0, HAlign.JUSTIFY: 0.0}


def draw_text_body(
    c: canvas_module.Canvas,
    text_body: TextBody,
    rect_pt: RectPt,
    font_registry: FontRegistry,
    math_renderer: LatexMathRenderer,
    warning_log: WarningLog,
    slide_index: int | None,
    page_index: int,
    overlay_items: list[OverlayItem],
    rotated_or_flipped: bool = False,
) -> None:
    """テキスト本体を矩形領域内へ描画する．

    引数:
        c (canvas_module.Canvas): 描画対象のキャンバス．
        text_body (TextBody): 描画対象のテキスト本体．
        rect_pt (RectPt): PDF座標系での図形の矩形（左下原点）．
        font_registry (FontRegistry): フォント解決用レジストリ．
        math_renderer (LatexMathRenderer): 数式レンダラー．
        warning_log (WarningLog): 警告記録先．
        slide_index (int | None): 対象スライド番号．
        page_index (int): 出力PDFにおけるページ番号（0始まり，数式合成用）．
        overlay_items (list[OverlayItem]): 数式PDFの合成予約リスト（追記される）．
        rotated_or_flipped (bool): 図形が回転・反転している場合True．
            この場合，数式のオーバーレイ合成は位置計算が複雑になるため
            サポート対象外とし，警告を出す．
    戻り値:
        なし．
    """

    inset_left_pt = text_body.inset_left_emu / 12700.0
    inset_top_pt = text_body.inset_top_emu / 12700.0
    inset_right_pt = text_body.inset_right_emu / 12700.0
    inset_bottom_pt = text_body.inset_bottom_emu / 12700.0

    available_width = max(1.0, rect_pt.width - inset_left_pt - inset_right_pt)
    lines = layout_text_body(text_body, available_width, font_registry, math_renderer, warning_log, slide_index)
    if not lines:
        return

    content_height = sum(line.space_before_pt + line.line_height_pt for line in lines)
    available_height = rect_pt.height - inset_top_pt - inset_bottom_pt

    if text_body.anchor == VAlign.TOP:
        content_top = rect_pt.height - inset_top_pt
    elif text_body.anchor == VAlign.MIDDLE:
        content_top = rect_pt.height - inset_top_pt - max(0.0, (available_height - content_height) / 2.0)
    else:
        content_top = inset_bottom_pt + content_height

    cursor_top = content_top
    for line in lines:
        cursor_top -= line.space_before_pt
        baseline_y_local = cursor_top - line.ascent_pt
        _draw_line(
            c,
            line,
            inset_left_pt,
            baseline_y_local,
            available_width,
            rect_pt,
            page_index,
            overlay_items,
            warning_log,
            slide_index,
            rotated_or_flipped,
            font_registry,
        )
        cursor_top -= line.line_height_pt


def _draw_text_at(c: canvas_module.Canvas, x: float, y: float, seg) -> None:
    """1つのテキストセグメントを座標(x, y)へ描画する（文字間隔`spc`を考慮）．"""

    if seg.char_spacing_pt:
        text_obj = c.beginText(x, y)
        text_obj.setFont(seg.reportlab_font, seg.size_pt)
        text_obj.setFillColorRGB(*seg.color.to_unit_tuple())
        text_obj.setCharSpace(seg.char_spacing_pt)
        text_obj.textOut(seg.text)
        c.drawText(text_obj)
    else:
        c.drawString(x, y, seg.text)


def _draw_line(
    c: canvas_module.Canvas,
    line: LineBox,
    inset_left_pt: float,
    baseline_y_local: float,
    available_width: float,
    rect_pt: RectPt,
    page_index: int,
    overlay_items: list[OverlayItem],
    warning_log: WarningLog,
    slide_index: int | None,
    rotated_or_flipped: bool,
    font_registry: FontRegistry,
) -> None:
    align_factor = _ALIGN_FACTOR[line.align]
    indented_width = max(0.0, available_width - line.indent_left_pt)
    x_local = (
        inset_left_pt
        + line.indent_left_pt
        + max(0.0, (indented_width - line.total_width_pt) * align_factor)
    )

    if line.bullet_char and line.segments:
        bullet_color = line.segments[0].color
        resolved = font_registry.resolve(line.bullet_font or "Arial", False, False, slide_index)
        bullet_font_name = font_registry.register_reportlab_font(resolved)
        bullet_x = rect_pt.x + inset_left_pt + line.indent_left_pt + line.bullet_offset_pt
        bullet_y = rect_pt.y + baseline_y_local
        c.setFont(bullet_font_name, line.bullet_size_pt)
        c.setFillColorRGB(*bullet_color.to_unit_tuple())
        c.setFillAlpha(line.segments[0].alpha)
        c.drawString(bullet_x, bullet_y, line.bullet_char)

    for seg in line.segments:
        if seg.kind == "text":
            abs_x = rect_pt.x + x_local
            abs_y = rect_pt.y + baseline_y_local
            c.setFont(seg.reportlab_font, seg.size_pt)
            c.setFillColorRGB(*seg.color.to_unit_tuple())
            c.setFillAlpha(seg.alpha)

            if seg.needs_faux_italic:
                c.saveState()
                c.translate(abs_x, abs_y)
                c.skew(12, 0)
                if seg.needs_faux_bold:
                    _draw_text_at(c, 0.3, 0, seg)
                _draw_text_at(c, 0, 0, seg)
                c.restoreState()
            else:
                if seg.needs_faux_bold:
                    _draw_text_at(c, abs_x + 0.3, abs_y, seg)
                _draw_text_at(c, abs_x, abs_y, seg)

            if seg.underline:
                underline_y = abs_y - seg.size_pt * 0.08
                c.setStrokeColorRGB(*seg.color.to_unit_tuple())
                c.setStrokeAlpha(seg.alpha)
                c.setLineWidth(max(0.5, seg.size_pt * 0.05))
                c.line(abs_x, underline_y, abs_x + seg.width_pt, underline_y)

        elif seg.kind == "math":
            if rotated_or_flipped:
                warning_log.add(
                    "unsupported_layout",
                    "回転・反転した図形内の数式描画は未対応のため省略しました．",
                    slide_index,
                )
            else:
                abs_x = rect_pt.x + x_local
                abs_y = rect_pt.y + baseline_y_local - seg.descent_pt
                overlay_items.append(
                    OverlayItem(
                        page_index=page_index,
                        pdf_bytes=seg.math_result.pdf_bytes,
                        x_pt=abs_x,
                        y_pt=abs_y,
                        width_pt=seg.width_pt,
                        height_pt=seg.math_result.total_height_pt,
                    )
                )

        x_local += seg.width_pt
