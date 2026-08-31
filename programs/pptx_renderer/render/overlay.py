"""ベクターPDF（数式・SVG由来）を，ReportLabで生成したベースページへ
後段で合成（オーバーレイ）するモジュール．

ReportLabは他のPDFをXObjectとして直接埋め込む機能を持たないため，
`pypdf`の`merge_transformed_page`を用いて，ページ単位で合成する．
これにより，数式・SVG画像を常にベクターのまま最終PDFへ埋め込める．
"""

from __future__ import annotations

import io
from dataclasses import dataclass

from pypdf import PdfReader, PdfWriter, Transformation


@dataclass
class OverlayItem:
    """1件の合成予約情報．

    属性:
        page_index (int): 合成先ページ番号（0始まり）．
        pdf_bytes (bytes): 合成するPDF（1ページ分）のバイト列．
        x_pt (float): 合成先の左下X座標（pt，ページ座標系）．
        y_pt (float): 合成先の左下Y座標（pt，ページ座標系）．
        width_pt (float): 合成先での幅（pt）．
        height_pt (float): 合成先での高さ（pt）．
    """

    page_index: int
    pdf_bytes: bytes
    x_pt: float
    y_pt: float
    width_pt: float
    height_pt: float
    rotation_deg: float = 0.0
    flip_h: bool = False
    flip_v: bool = False


def apply_overlays(base_pdf_bytes: bytes, overlay_items: list[OverlayItem]) -> bytes:
    """ベースPDFに対して，全ての合成予約を適用する．

    引数:
        base_pdf_bytes (bytes): ReportLabで生成したベースPDFのバイト列．
        overlay_items (list[OverlayItem]): 合成予約のリスト．
    戻り値:
        bytes: 合成後のPDFのバイト列．
    """

    if not overlay_items:
        return base_pdf_bytes

    reader = PdfReader(io.BytesIO(base_pdf_bytes))
    writer = PdfWriter()
    for pg in reader.pages:
        writer.add_page(pg)

    for item in overlay_items:
        overlay_reader = PdfReader(io.BytesIO(item.pdf_bytes))
        overlay_page = overlay_reader.pages[0]
        src_w = float(overlay_page.mediabox.width)
        src_h = float(overlay_page.mediabox.height)
        if src_w == 0 or src_h == 0:
            continue
        sx = item.width_pt / src_w
        sy = item.height_pt / src_h
        cx, cy = item.width_pt / 2.0, item.height_pt / 2.0

        transform = Transformation().scale(sx, sy).translate(-cx, -cy)
        if item.flip_h or item.flip_v:
            transform = transform.scale(-1 if item.flip_h else 1, -1 if item.flip_v else 1)
        if item.rotation_deg:
            # PPTXの回転角は時計回りが正，pypdfの`rotate`は反時計回りが正のため符号を反転する．
            transform = transform.rotate(-item.rotation_deg)
        transform = transform.translate(item.x_pt + cx, item.y_pt + cy)
        writer.pages[item.page_index].merge_transformed_page(overlay_page, transform)

    out = io.BytesIO()
    writer.write(out)
    return out.getvalue()
