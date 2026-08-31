"""1スライド分のXML（`ppt/slides/slideN.xml`）をSlide IRへ変換するパーサー．

背景・図形ツリー（`p:spTree`）を解析し，PowerPointの描画順序を保持したまま
`pptx_renderer.ir.Slide`を構築する．
"""

from __future__ import annotations

from lxml import etree

from pptx_renderer.ir import Background, Fill, Slide
from pptx_renderer.parser.common import a, iter_effective_children, p, resolve_color_element
from pptx_renderer.parser.package import PptxPackage
from pptx_renderer.parser.shape_parser import parse_cxn_sp, parse_grp_sp, parse_pic, parse_sp
from pptx_renderer.parser.table_parser import parse_graphic_frame
from pptx_renderer.parser.theme import Theme, parse_theme
from pptx_renderer.warnings_log import WarningLog

def _parse_background(bg_el: etree._Element | None, theme: Theme) -> Fill | None:
    if bg_el is None:
        return None
    bg_pr = bg_el.find(p("bgPr"))
    if bg_pr is not None:
        solid = bg_pr.find(a("solidFill"))
        if solid is not None:
            color_el = solid.find(a("srgbClr"))
            if color_el is None:
                color_el = solid.find(a("schemeClr"))
            if color_el is not None:
                color, alpha = resolve_color_element(color_el, theme)
                return Fill(kind="solid", color=color, alpha=alpha)
        return None

    bg_ref = bg_el.find(p("bgRef"))
    if bg_ref is not None:
        color_el = bg_ref.find(a("schemeClr"))
        if color_el is None:
            color_el = bg_ref.find(a("srgbClr"))
        if color_el is not None:
            color, alpha = resolve_color_element(color_el, theme)
            return Fill(kind="solid", color=color, alpha=alpha)

    return None


def _resolve_background(
    package: PptxPackage, slide_part: str, slide_root: etree._Element, theme: Theme
) -> Background:
    fill = _parse_background(slide_root.find(p("cSld") + "/" + p("bg")), theme)
    if fill is None:
        layout_part = package.slide_layout_for_slide(slide_part)
        if layout_part is not None:
            layout_root = package.read_xml(layout_part)
            fill = _parse_background(layout_root.find(p("cSld") + "/" + p("bg")), theme)
            if fill is None:
                master_part = package.slide_master_for_layout(layout_part)
                if master_part is not None:
                    master_root = package.read_xml(master_part)
                    fill = _parse_background(master_root.find(p("cSld") + "/" + p("bg")), theme)

    if fill is None or fill.color is None:
        from pptx_renderer.ir import WHITE

        fill = Fill(kind="solid", color=WHITE, alpha=1.0)

    return Background(fill=fill)


def _is_placeholder(shape_el: etree._Element) -> bool:
    return shape_el.find(f".//{p('ph')}") is not None


def _parse_non_placeholder_shapes(
    root: etree._Element,
    theme: Theme,
    package: PptxPackage,
    part_name: str,
    warning_log: WarningLog,
    slide_index: int,
) -> list:
    """スライドレイアウト／マスターの装飾図形（プレースホルダを除く）を解析する．

    スライド自身のプレースホルダによって上書きされない，背景の帯や
    ロゴ画像等の固定要素をスライドマスター・レイアウトから引き継ぐために使用する．
    """

    from pptx_renderer.parser.shape_parser import parse_cxn_sp, parse_grp_sp, parse_pic, parse_sp
    from pptx_renderer.parser.table_parser import parse_graphic_frame

    sp_tree = root.find(f"{p('cSld')}/{p('spTree')}")
    if sp_tree is None:
        return []

    shapes = []
    for child in iter_effective_children(sp_tree):
        local = etree.QName(child).localname
        if local in ("nvGrpSpPr", "grpSpPr"):
            continue
        if _is_placeholder(child):
            continue
        if local == "sp":
            shapes.append(parse_sp(child, theme, warning_log, slide_index))
        elif local == "pic":
            pic = parse_pic(child, theme, package, part_name, warning_log, slide_index)
            if pic is not None:
                shapes.append(pic)
        elif local == "cxnSp":
            shapes.append(parse_cxn_sp(child, theme, warning_log, slide_index))
        elif local == "grpSp":
            shapes.append(parse_grp_sp(child, theme, package, part_name, warning_log, slide_index))
        elif local == "graphicFrame":
            table = parse_graphic_frame(child, theme, package, warning_log, slide_index)
            if table is not None:
                shapes.append(table)
    return shapes


def parse_slide(
    package: PptxPackage,
    slide_part: str,
    slide_index: int,
    slide_width_emu: float,
    slide_height_emu: float,
    warning_log: WarningLog,
) -> Slide:
    """1スライド分のXMLをSlide IRへ変換する．

    引数:
        package (PptxPackage): PPTXパッケージ．
        slide_part (str): スライドのパート名（例: "ppt/slides/slide1.xml"）．
        slide_index (int): スライド番号（1始まり，警告表示用）．
        slide_width_emu (float): スライド幅（EMU）．
        slide_height_emu (float): スライド高さ（EMU）．
        warning_log (WarningLog): 警告記録先．
    戻り値:
        Slide: 変換結果．
    """

    slide_root = package.read_xml(slide_part)

    theme_part = package.theme_for_slide(slide_part)
    theme = parse_theme(package.read_xml(theme_part) if theme_part else None)

    background = _resolve_background(package, slide_part, slide_root, theme)

    shapes = []
    layout_part = package.slide_layout_for_slide(slide_part)
    if layout_part is not None:
        master_part = package.slide_master_for_layout(layout_part)
        if master_part is not None:
            shapes.extend(
                _parse_non_placeholder_shapes(
                    package.read_xml(master_part), theme, package, master_part, warning_log, slide_index
                )
            )
        shapes.extend(
            _parse_non_placeholder_shapes(
                package.read_xml(layout_part), theme, package, layout_part, warning_log, slide_index
            )
        )

    sp_tree = slide_root.find(f"{p('cSld')}/{p('spTree')}")
    if sp_tree is not None:
        for child in iter_effective_children(sp_tree):
            local = etree.QName(child).localname
            if local == "sp":
                shapes.append(parse_sp(child, theme, warning_log, slide_index, package, slide_part))
            elif local == "pic":
                pic = parse_pic(child, theme, package, slide_part, warning_log, slide_index)
                if pic is not None:
                    shapes.append(pic)
            elif local == "cxnSp":
                shapes.append(parse_cxn_sp(child, theme, warning_log, slide_index))
            elif local == "grpSp":
                shapes.append(parse_grp_sp(child, theme, package, slide_part, warning_log, slide_index))
            elif local == "graphicFrame":
                table = parse_graphic_frame(child, theme, package, warning_log, slide_index)
                if table is not None:
                    shapes.append(table)
            elif local in ("nvGrpSpPr", "grpSpPr"):
                continue
            else:
                warning_log.unsupported_element(f"p:{local}", slide_index)

    return Slide(
        index=slide_index,
        width_emu=slide_width_emu,
        height_emu=slide_height_emu,
        background=background,
        shapes=shapes,
    )
