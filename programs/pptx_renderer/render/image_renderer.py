"""画像（PNG, JPEG, SVG）の描画を行うモジュール．

ラスタ画像（PNG/JPEG）はReportLabの`drawImage`で直接描画する．
SVG画像は，ラスタ画像へ変換せず，`cairosvg`でベクターPDFへ変換した上で
`pptx_renderer.render.overlay`を通じてページへ合成する（真のベクター画像
として埋め込むため）．

トリミング（crop）はPillowで画素単位に切り出してから描画する．
無意味な低解像度化は行わない（元の解像度のまま埋め込む）．
"""

from __future__ import annotations

import io

import cairosvg
from PIL import Image
from reportlab.lib.utils import ImageReader
from reportlab.pdfgen import canvas as canvas_module

from pptx_renderer.ir import PictureShape
from pptx_renderer.render.overlay import OverlayItem
from pptx_renderer.render.shape_renderer import apply_shape_transform
from pptx_renderer.units import RectPt
from pptx_renderer.warnings_log import WarningLog


def _apply_crop(image_bytes: bytes, crop_left: float, crop_top: float, crop_right: float, crop_bottom: float) -> bytes:
    if crop_left <= 0 and crop_top <= 0 and crop_right <= 0 and crop_bottom <= 0:
        return image_bytes

    img = Image.open(io.BytesIO(image_bytes))
    w, h = img.size
    box = (
        int(w * crop_left),
        int(h * crop_top),
        int(w * (1.0 - crop_right)),
        int(h * (1.0 - crop_bottom)),
    )
    cropped = img.crop(box)
    out = io.BytesIO()
    cropped.save(out, format=img.format or "PNG")
    return out.getvalue()


def draw_picture(
    c: canvas_module.Canvas,
    shape: PictureShape,
    rect_pt: RectPt,
    page_index: int,
    overlay_items: list[OverlayItem],
    warning_log: WarningLog,
    slide_index: int | None,
) -> None:
    """`PictureShape`をPDFへ描画する．

    引数:
        c (canvas_module.Canvas): 描画対象のキャンバス．
        shape (PictureShape): 描画対象の画像図形．
        rect_pt (RectPt): PDF座標系での配置矩形．
        page_index (int): 出力PDFにおけるページ番号（0始まり）．
        overlay_items (list[OverlayItem]): SVG合成予約リスト（追記される）．
        warning_log (WarningLog): 警告記録先．
        slide_index (int | None): 対象スライド番号．
    戻り値:
        なし．
    """

    if shape.image_format == "svg":
        try:
            pdf_bytes = cairosvg.svg2pdf(bytestring=shape.image_bytes)
        except Exception as exc:  # noqa: BLE001
            warning_log.add("image_render_error", f"SVG画像の変換に失敗しました: {exc}", slide_index)
            return
        overlay_items.append(
            OverlayItem(
                page_index=page_index,
                pdf_bytes=pdf_bytes,
                x_pt=rect_pt.x,
                y_pt=rect_pt.y,
                width_pt=rect_pt.width,
                height_pt=rect_pt.height,
                rotation_deg=shape.rotation,
                flip_h=shape.flip_h,
                flip_v=shape.flip_v,
            )
        )
        return

    try:
        image_bytes = _apply_crop(
            shape.image_bytes, shape.crop_left, shape.crop_top, shape.crop_right, shape.crop_bottom
        )
        reader = ImageReader(io.BytesIO(image_bytes))
    except Exception as exc:  # noqa: BLE001
        warning_log.add("image_render_error", f"画像の読み込みに失敗しました: {exc}", slide_index)
        return

    c.saveState()
    apply_shape_transform(c, rect_pt, shape.rotation, shape.flip_h, shape.flip_v)
    c.setFillAlpha(shape.alpha)
    c.drawImage(reader, 0, 0, width=rect_pt.width, height=rect_pt.height, mask="auto", preserveAspectRatio=False)
    c.restoreState()
