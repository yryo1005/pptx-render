"""Slide IRからPDFを生成するメインのレンダラー．

スライドごとに背景→画像→図形→テキストの順（PowerPointの描画順序）で
ReportLabのキャンバスへ描画し，最後に数式・SVGのベクターPDFを
`pptx_renderer.render.overlay`で合成して最終PDFを出力する．

グループ化された図形（`p:grpSp`）は，描画前に子図形の絶対座標へ
変換（フラット化）してから描画する．これにより，各図形の描画処理を
グループの有無によらず統一的に扱える．
"""

from __future__ import annotations

import io
import math
import os
from concurrent.futures import ThreadPoolExecutor

from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas as canvas_module

from pptx_renderer.fonts.registry import FontRegistry, default_registry
from pptx_renderer.ir import AutoShape, GroupShape, PictureShape, Presentation, ShapeIR, Slide, TableShape
from pptx_renderer.math.latex_render import LatexMathRenderer
from pptx_renderer.render.image_renderer import draw_picture
from pptx_renderer.render.overlay import OverlayItem, apply_overlays
from pptx_renderer.render.shape_renderer import draw_auto_shape
from pptx_renderer.render.table_renderer import draw_table
from pptx_renderer.render.text_renderer import draw_text_body
from pptx_renderer.units import CoordinateTransformer, RectEMU
from pptx_renderer.warnings_log import WarningLog, default_log


def _rotate_point(px: float, py: float, cx: float, cy: float, rotation_deg: float) -> tuple[float, float]:
    """点(px, py)を中心(cx, cy)まわりに，PPTXの角度規約（時計回り，Y軸下向き）で回転する．"""

    rad = math.radians(rotation_deg)
    dx, dy = px - cx, py - cy
    cos_r, sin_r = math.cos(rad), math.sin(rad)
    return cx + dx * cos_r - dy * sin_r, cy + dx * sin_r + dy * cos_r


class _FlatShape:
    """グループ座標変換を適用した後の，描画用フラット化図形．"""

    __slots__ = ("shape", "abs_rect", "abs_rotation")

    def __init__(self, shape: ShapeIR, abs_rect: RectEMU, abs_rotation: float) -> None:
        self.shape = shape
        self.abs_rect = abs_rect
        self.abs_rotation = abs_rotation


def _flatten_shapes(
    shapes: list[ShapeIR],
    parent_rect: RectEMU | None = None,
    parent_rotation: float = 0.0,
) -> list[_FlatShape]:
    """図形リストを，グループを展開したフラットなリストへ変換する．

    引数:
        shapes (list[ShapeIR]): 変換対象の図形リスト（グループを含みうる）．
        parent_rect (RectEMU | None): 親グループの絶対矩形（トップレベルではNone）．
        parent_rotation (float): 親グループまでの累積回転角（度）．
    戻り値:
        list[_FlatShape]: 描画順を保ったフラットな図形リスト．
    """

    result: list[_FlatShape] = []
    for shape in shapes:
        if isinstance(shape, GroupShape):
            group_abs_rect, group_abs_rotation = _compose_rect(shape.rect, shape.rotation, parent_rect, parent_rotation)
            ch_off_x, ch_off_y = shape.child_offset_emu
            ch_ext_x, ch_ext_y = shape.child_extent_emu
            ch_ext_x = ch_ext_x or 1.0
            ch_ext_y = ch_ext_y or 1.0

            scale_x = group_abs_rect.cx / ch_ext_x
            scale_y = group_abs_rect.cy / ch_ext_y

            remapped_children = []
            for child in shape.children:
                child_rect_in_group_local = RectEMU(
                    x=(child.rect.x - ch_off_x) * scale_x,
                    y=(child.rect.y - ch_off_y) * scale_y,
                    cx=child.rect.cx * scale_x,
                    cy=child.rect.cy * scale_y,
                )
                remapped_children.append(_with_rect(child, child_rect_in_group_local))

            result.extend(_flatten_shapes(remapped_children, group_abs_rect, group_abs_rotation))
        else:
            abs_rect, abs_rotation = _compose_rect(shape.rect, shape.rotation, parent_rect, parent_rotation)
            result.append(_FlatShape(shape, abs_rect, abs_rotation))
    return result


def _compose_rect(
    rect: RectEMU, rotation: float, parent_rect: RectEMU | None, parent_rotation: float
) -> tuple[RectEMU, float]:
    if parent_rect is None:
        return rect, rotation

    center_x, center_y = rect.center_x, rect.center_y
    parent_center_x, parent_center_y = parent_rect.center_x, parent_rect.center_y
    rot_x, rot_y = _rotate_point(center_x, center_y, parent_center_x, parent_center_y, parent_rotation)

    abs_rect = RectEMU(x=rot_x - rect.cx / 2.0, y=rot_y - rect.cy / 2.0, cx=rect.cx, cy=rect.cy)
    abs_rotation = rotation + parent_rotation
    return abs_rect, abs_rotation


def _with_rect(shape: ShapeIR, new_rect: RectEMU) -> ShapeIR:
    """図形の`rect`のみを差し替えた同種オブジェクトを返す（浅いコピー）．"""

    import copy

    copied = copy.copy(shape)
    copied.rect = new_rect
    return copied


class PdfRenderer:
    """Slide IRを入力としてPDFを生成するレンダラー本体．"""

    def __init__(
        self,
        warning_log: WarningLog | None = None,
        font_registry: FontRegistry | None = None,
        use_math_disk_cache: bool = True,
    ) -> None:
        """コンストラクタ．

        引数:
            warning_log (WarningLog | None): 警告記録先．Noneの場合は共有ロガーを使用する．
            font_registry (FontRegistry | None): フォントレジストリ．Noneの場合は共有インスタンスを使用する．
            use_math_disk_cache (bool): 数式レンダリング結果のディスクキャッシュを使用するか．
        戻り値:
            なし．
        """

        self._warning_log = warning_log or default_log
        self._font_registry = font_registry or default_registry
        self._math_renderer = LatexMathRenderer(self._warning_log, use_disk_cache=use_math_disk_cache)

    def render(
        self,
        presentation: Presentation,
        output_path: str,
        render_pages: list[int] | None = None,
        max_workers: int | None = None,
    ) -> None:
        """プレゼンテーション全体をPDFへ描画し，ファイルへ書き出す．

        各スライドは独立した1ページ分のPDFとして並列にレンダリングし，
        最後に結合する．数式レンダリング（`xelatex`/`dvisvgm`の起動）が
        支配的なコストであり，これらはサブプロセス待機中にGILを解放する
        I/Oバウンドな処理であるため，スレッド並列でも実行時間の短縮が
        見込める．

        引数:
            presentation (Presentation): 描画対象のSlide IR．
            output_path (str): 出力PDFファイルのパス．
            render_pages (list[int] | None): 描画対象のスライド番号（1始まり）のリスト．
                Noneの場合は全スライドを描画する．
            max_workers (int | None): 並列描画に使用するワーカー数．
                Noneの場合はCPUコア数とスライド数から自動決定する．
        戻り値:
            なし．
        """

        transformer = CoordinateTransformer(presentation.width_emu, presentation.height_emu)
        target_slides = [
            slide for slide in presentation.slides if render_pages is None or slide.index in render_pages
        ]

        if max_workers is None:
            max_workers = min(8, os.cpu_count() or 1, max(1, len(target_slides)))

        if max_workers <= 1 or len(target_slides) <= 1:
            page_pdf_bytes = [self._render_single_page(slide, transformer) for slide in target_slides]
        else:
            with ThreadPoolExecutor(max_workers=max_workers) as executor:
                page_pdf_bytes = list(
                    executor.map(lambda slide: self._render_single_page(slide, transformer), target_slides)
                )

        writer = PdfWriter()
        for pdf_bytes in page_pdf_bytes:
            reader = PdfReader(io.BytesIO(pdf_bytes))
            writer.add_page(reader.pages[0])
        with open(output_path, "wb") as f:
            writer.write(f)

    def _render_single_page(self, slide: Slide, transformer: CoordinateTransformer) -> bytes:
        """1スライド分を独立したPDF（1ページ）として描画する．

        引数:
            slide (Slide): 描画対象のスライド．
            transformer (CoordinateTransformer): EMU→pt座標変換器．
        戻り値:
            bytes: 数式・SVGの合成まで完了した，1ページ分のPDFバイト列．
        """

        buffer = io.BytesIO()
        c = canvas_module.Canvas(buffer, pagesize=(transformer.page_width_pt, transformer.page_height_pt))
        overlay_items: list[OverlayItem] = []
        self._render_slide(c, slide, transformer, 0, overlay_items)
        c.showPage()
        c.save()
        return apply_overlays(buffer.getvalue(), overlay_items)

    def _render_slide(
        self,
        c: canvas_module.Canvas,
        slide: Slide,
        transformer: CoordinateTransformer,
        page_index: int,
        overlay_items: list[OverlayItem],
    ) -> None:
        if slide.background.fill.kind == "solid" and slide.background.fill.color is not None:
            c.setFillColorRGB(*slide.background.fill.color.to_unit_tuple())
            c.setFillAlpha(slide.background.fill.alpha)
            c.rect(0, 0, transformer.page_width_pt, transformer.page_height_pt, fill=1, stroke=0)

        flat_shapes = _flatten_shapes(slide.shapes)
        for flat in flat_shapes:
            self._render_shape(c, flat, transformer, page_index, overlay_items, slide.index)

    def _render_shape(
        self,
        c: canvas_module.Canvas,
        flat: _FlatShape,
        transformer: CoordinateTransformer,
        page_index: int,
        overlay_items: list[OverlayItem],
        slide_index: int,
    ) -> None:
        shape = flat.shape
        rect_pt = transformer.rect_to_pdf(flat.abs_rect)

        if isinstance(shape, PictureShape):
            shape_for_draw = shape
            shape_for_draw.rotation = flat.abs_rotation
            draw_picture(c, shape_for_draw, rect_pt, page_index, overlay_items, self._warning_log, slide_index)
            return

        if isinstance(shape, TableShape):
            draw_table(
                c,
                shape,
                transformer,
                (flat.abs_rect.x, flat.abs_rect.y),
                self._font_registry,
                self._math_renderer,
                page_index,
                overlay_items,
                self._warning_log,
                slide_index,
            )
            return

        if isinstance(shape, AutoShape):
            shape_for_draw = shape
            shape_for_draw.rotation = flat.abs_rotation
            draw_auto_shape(c, shape_for_draw, rect_pt, self._warning_log, slide_index)
            if shape.text_body is not None:
                rotated_or_flipped = bool(flat.abs_rotation) or shape.flip_h or shape.flip_v
                draw_text_body(
                    c,
                    shape.text_body,
                    rect_pt,
                    self._font_registry,
                    self._math_renderer,
                    self._warning_log,
                    slide_index,
                    page_index,
                    overlay_items,
                    rotated_or_flipped=rotated_or_flipped,
                )
            return

        self._warning_log.unsupported_element(type(shape).__name__, slide_index)
