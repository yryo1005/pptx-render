"""座標変換モジュール（`pptx_renderer.units`）の単体テスト．"""

from __future__ import annotations

from pptx_renderer.units import CoordinateTransformer, RectEMU, emu_to_pt, pt_to_emu


def test_emu_to_pt_roundtrip() -> None:
    assert abs(emu_to_pt(pt_to_emu(72.0)) - 72.0) < 1e-9


def test_emu_to_pt_known_value() -> None:
    # 1インチ = 914400 EMU = 72pt
    assert abs(emu_to_pt(914400) - 72.0) < 1e-9


def test_coordinate_transformer_page_size() -> None:
    transformer = CoordinateTransformer(slide_width_emu=12192000, slide_height_emu=6858000)
    assert abs(transformer.page_width_pt - 960.0) < 1e-6
    assert abs(transformer.page_height_pt - 540.0) < 1e-6


def test_rect_to_pdf_flips_y_axis() -> None:
    transformer = CoordinateTransformer(slide_width_emu=914400 * 10, slide_height_emu=914400 * 10)
    rect = RectEMU(x=0, y=0, cx=914400, cy=914400)
    rect_pt = transformer.rect_to_pdf(rect)
    # スライド左上のEMU矩形は，PDF座標系ではページ上端（y = height - 1inch）に位置する．
    assert abs(rect_pt.x - 0.0) < 1e-6
    assert abs(rect_pt.y - (720.0 - 72.0)) < 1e-6
    assert abs(rect_pt.width - 72.0) < 1e-6
    assert abs(rect_pt.height - 72.0) < 1e-6
