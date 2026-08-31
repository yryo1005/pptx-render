"""個々の図形（`p:sp`, `p:pic`, `p:cxnSp`, `p:grpSp`）をSlide IRへ変換するパーサー．"""

from __future__ import annotations

from lxml import etree

from pptx_renderer.ir import AutoShape, Fill, GroupShape, PictureShape, ShapeIR, ShapeStyle
from pptx_renderer.parser.common import a, iter_effective_children, p, parse_fill, parse_line, parse_xfrm, r
from pptx_renderer.parser.custgeom_parser import parse_cust_geom
from pptx_renderer.parser.package import PptxPackage
from pptx_renderer.parser.placeholder import resolve_placeholder_level_defaults, resolve_placeholder_rect
from pptx_renderer.parser.text_parser import parse_text_body
from pptx_renderer.parser.theme import Theme
from pptx_renderer.warnings_log import WarningLog


def _shape_id_name(nv_pr_el: etree._Element) -> tuple[str, str]:
    c_nv_pr = nv_pr_el.find(f"{p('cNvPr')}")
    if c_nv_pr is None:
        return "0", ""
    return c_nv_pr.get("id", "0"), c_nv_pr.get("name", "")


_PLACEHOLDER_DEFAULT_SIZE = {
    "title": 44.0,
    "ctrTitle": 44.0,
    "subTitle": 28.0,
}


def _default_text_size(sp_el: etree._Element) -> float:
    ph = sp_el.find(f"{p('nvSpPr')}/{p('nvPr')}/{p('ph')}")
    if ph is not None:
        return _PLACEHOLDER_DEFAULT_SIZE.get(ph.get("type", ""), 18.0)
    return 18.0


def parse_sp(
    sp_el: etree._Element,
    theme: Theme,
    warning_log: WarningLog,
    slide_index: int | None,
    package: PptxPackage | None = None,
    slide_part: str | None = None,
) -> AutoShape:
    """`p:sp`（自由図形・テキストボックス・プレースホルダ）を`AutoShape`へ変換する．

    自身の`a:xfrm`が省略されているプレースホルダについては，スライド
    レイアウト・スライドマスターから位置・サイズを継承解決する．

    引数:
        sp_el (etree._Element): `p:sp`要素．
        theme (Theme): テーマ情報．
        warning_log (WarningLog): 警告記録先．
        slide_index (int | None): 対象スライド番号．
        package (PptxPackage | None): プレースホルダ継承解決用のパッケージ．
        slide_part (str | None): 対象スライドのパート名．
    戻り値:
        AutoShape: 変換結果．
    """

    shape_id, name = _shape_id_name(sp_el.find(p("nvSpPr")))
    sp_pr = sp_el.find(p("spPr"))
    xfrm = sp_pr.find(a("xfrm")) if sp_pr is not None else None
    rect, rotation, flip_h, flip_v = parse_xfrm(xfrm)

    default_color = None
    default_bold = False
    default_size = _default_text_size(sp_el)
    level_defaults_fn = None
    ph = sp_el.find(f"{p('nvSpPr')}/{p('nvPr')}/{p('ph')}")
    if ph is not None and package is not None and slide_part is not None:
        if rect.cx == 0 and rect.cy == 0:
            inherited = resolve_placeholder_rect(package, slide_part, ph.get("type"), ph.get("idx"))
            if inherited is not None:
                rect = inherited

        ph_type, ph_idx = ph.get("type"), ph.get("idx")

        def level_defaults_fn(level: int, _ph_type=ph_type, _ph_idx=ph_idx) -> dict:
            return resolve_placeholder_level_defaults(package, slide_part, _ph_type, _ph_idx, level, theme)

        defaults0 = level_defaults_fn(0)
        default_color = defaults0.get("color")
        default_bold = defaults0.get("bold", False)
        default_size = defaults0.get("size_pt", default_size)

    preset = "rect"
    adjustments: dict[str, float] = {}
    custom_paths = []
    if sp_pr is not None:
        prst_geom = sp_pr.find(a("prstGeom"))
        cust_geom = sp_pr.find(a("custGeom"))
        if prst_geom is not None:
            preset = prst_geom.get("prst", "rect")
            for gd in prst_geom.findall(f"{a('avLst')}/{a('gd')}"):
                gd_name = gd.get("name")
                gd_val = gd.get("fmla", "")
                if gd_val.startswith("val "):
                    try:
                        adjustments[gd_name] = float(gd_val.split()[1])
                    except (IndexError, ValueError):
                        pass
        elif cust_geom is not None:
            preset = "custGeom"
            custom_paths = parse_cust_geom(cust_geom)
            if not custom_paths:
                warning_log.unsupported_shape("custGeom（パス無し）を矩形として近似しました．", slide_index)

    style_el = sp_el.find(p("style"))
    style = ShapeStyle(
        fill=parse_fill(sp_pr, theme, warning_log, slide_index),
        stroke=parse_line(sp_pr, theme, warning_log, slide_index, style_el),
    )

    tx_body_el = sp_el.find(p("txBody"))
    text_body = parse_text_body(
        tx_body_el,
        theme,
        warning_log,
        slide_index,
        default_size_pt=default_size,
        default_color=default_color,
        default_bold=default_bold,
        level_defaults_fn=level_defaults_fn,
    )

    return AutoShape(
        id=shape_id,
        name=name,
        rect=rect,
        preset=preset,
        style=style,
        text_body=text_body,
        rotation=rotation,
        flip_h=flip_h,
        flip_v=flip_v,
        adjustments=adjustments,
        custom_paths=custom_paths,
    )


def parse_cxn_sp(
    cxn_el: etree._Element, theme: Theme, warning_log: WarningLog, slide_index: int | None
) -> AutoShape:
    """`p:cxnSp`（コネクタ・線）を`AutoShape`へ変換する．"""

    shape_id, name = _shape_id_name(cxn_el.find(p("nvCxnSpPr")))
    sp_pr = cxn_el.find(p("spPr"))
    xfrm = sp_pr.find(a("xfrm")) if sp_pr is not None else None
    rect, rotation, flip_h, flip_v = parse_xfrm(xfrm)

    nv_cxn_pr = cxn_el.find(p("nvCxnSpPr"))
    st_cxn = nv_cxn_pr.find(f"{p('cNvCxnSpPr')}/{a('stCxn')}") if nv_cxn_pr is not None else None
    end_cxn = nv_cxn_pr.find(f"{p('cNvCxnSpPr')}/{a('endCxn')}") if nv_cxn_pr is not None else None
    start_connect_idx = int(st_cxn.get("idx")) if st_cxn is not None else None
    end_connect_idx = int(end_cxn.get("idx")) if end_cxn is not None else None

    preset = "line"
    if sp_pr is not None:
        prst_geom = sp_pr.find(a("prstGeom"))
        if prst_geom is not None:
            preset = prst_geom.get("prst", "line")

    style_el = cxn_el.find(p("style"))
    style = ShapeStyle(
        fill=parse_fill(sp_pr, theme, warning_log, slide_index),
        stroke=parse_line(sp_pr, theme, warning_log, slide_index, style_el),
    )

    return AutoShape(
        id=shape_id,
        name=name,
        rect=rect,
        preset=preset,
        style=style,
        text_body=None,
        rotation=rotation,
        flip_h=flip_h,
        flip_v=flip_v,
        start_connect_idx=start_connect_idx,
        end_connect_idx=end_connect_idx,
    )


_ASVG_NS = "http://schemas.microsoft.com/office/drawing/2016/SVG/main"


def _image_format_from_part_name(part_name: str) -> str:
    ext = part_name.rsplit(".", 1)[-1].lower()
    return {"jpg": "jpeg"}.get(ext, ext)


def parse_pic(
    pic_el: etree._Element,
    theme: Theme,
    package: PptxPackage,
    slide_part: str,
    warning_log: WarningLog,
    slide_index: int | None,
) -> PictureShape | None:
    """`p:pic`（画像）を`PictureShape`へ変換する．

    SVG拡張（`asvg:svgBlip`）が存在する場合はラスタ画像へ変換せず，
    SVGをそのままベクター画像として保持する．

    引数:
        pic_el (etree._Element): `p:pic`要素．
        theme (Theme): テーマ情報．
        package (PptxPackage): 画像バイト列解決用のパッケージ．
        slide_part (str): 画像を参照しているスライドのパート名．
        warning_log (WarningLog): 警告記録先．
        slide_index (int | None): 対象スライド番号．
    戻り値:
        PictureShape | None: 変換結果．画像を解決できない場合はNone．
    """

    shape_id, name = _shape_id_name(pic_el.find(p("nvPicPr")))
    blip_fill = pic_el.find(p("blipFill"))
    sp_pr = pic_el.find(p("spPr"))
    xfrm = sp_pr.find(a("xfrm")) if sp_pr is not None else None
    rect, rotation, flip_h, flip_v = parse_xfrm(xfrm)

    if blip_fill is None:
        warning_log.add("unsupported_element", "p:pic にblipFillが存在しません．", slide_index)
        return None

    blip = blip_fill.find(a("blip"))
    rid = None
    if blip is not None:
        svg_blip = blip.find(f"{a('extLst')}/{a('ext')}/{{{_ASVG_NS}}}svgBlip")
        if svg_blip is not None:
            rid = svg_blip.get(r("embed"))
        elif blip.get(r("embed")):
            rid = blip.get(r("embed"))

    if rid is None:
        warning_log.add("unsupported_element", "p:pic の画像参照（r:embed）を解決できませんでした．", slide_index)
        return None

    target = package.resolve_rid(slide_part, rid)
    if target is None or not package.exists(target):
        warning_log.add("unsupported_element", f"画像パート {target} が見つかりません．", slide_index)
        return None

    image_bytes = package.read_bytes(target)
    image_format = _image_format_from_part_name(target)

    crop_left = crop_top = crop_right = crop_bottom = 0.0
    src_rect = blip_fill.find(a("srcRect"))
    if src_rect is not None:
        crop_left = int(src_rect.get("l", "0")) / 100000.0
        crop_top = int(src_rect.get("t", "0")) / 100000.0
        crop_right = int(src_rect.get("r", "0")) / 100000.0
        crop_bottom = int(src_rect.get("b", "0")) / 100000.0

    alpha = 1.0
    if blip is not None:
        alpha_fix = blip.find(a("alphaModFix"))
        if alpha_fix is not None and alpha_fix.get("amt"):
            alpha = int(alpha_fix.get("amt")) / 100000.0

    style = ShapeStyle(
        fill=Fill(kind="none"),
        stroke=parse_line(sp_pr, theme, warning_log, slide_index),
    )

    return PictureShape(
        id=shape_id,
        name=name,
        rect=rect,
        image_bytes=image_bytes,
        image_format=image_format,
        crop_left=crop_left,
        crop_top=crop_top,
        crop_right=crop_right,
        crop_bottom=crop_bottom,
        alpha=alpha,
        style=style,
        rotation=rotation,
        flip_h=flip_h,
        flip_v=flip_v,
    )


def parse_grp_sp(
    grp_el: etree._Element,
    theme: Theme,
    package: PptxPackage,
    slide_part: str,
    warning_log: WarningLog,
    slide_index: int | None,
) -> GroupShape:
    """`p:grpSp`（グループ化された図形）を`GroupShape`へ変換する．

    子図形の座標は，グループ座標系（`chOff`/`chExt`）のまま保持し，
    実際の座標変換はレンダリング時にまとめて行う．

    引数:
        grp_el (etree._Element): `p:grpSp`要素．
        theme (Theme): テーマ情報．
        package (PptxPackage): 画像解決用のパッケージ．
        slide_part (str): スライドのパート名．
        warning_log (WarningLog): 警告記録先．
        slide_index (int | None): 対象スライド番号．
    戻り値:
        GroupShape: 変換結果．
    """

    shape_id, name = _shape_id_name(grp_el.find(p("nvGrpSpPr")))
    grp_sp_pr = grp_el.find(p("grpSpPr"))
    xfrm = grp_sp_pr.find(a("xfrm")) if grp_sp_pr is not None else None
    rect, rotation, flip_h, flip_v = parse_xfrm(xfrm)

    ch_off = (0.0, 0.0)
    ch_ext = (rect.cx if rect.cx else 1.0, rect.cy if rect.cy else 1.0)
    if xfrm is not None:
        ch_off_el = xfrm.find(a("chOff"))
        ch_ext_el = xfrm.find(a("chExt"))
        if ch_off_el is not None:
            ch_off = (float(ch_off_el.get("x")), float(ch_off_el.get("y")))
        if ch_ext_el is not None:
            ch_ext = (float(ch_ext_el.get("cx")), float(ch_ext_el.get("cy")))

    children: list[ShapeIR] = []
    from pptx_renderer.parser.table_parser import parse_graphic_frame  # 遅延importで循環参照を回避

    for child in iter_effective_children(grp_el):
        local = etree.QName(child).localname
        if local == "sp":
            children.append(parse_sp(child, theme, warning_log, slide_index, package, slide_part))
        elif local == "pic":
            pic = parse_pic(child, theme, package, slide_part, warning_log, slide_index)
            if pic is not None:
                children.append(pic)
        elif local == "cxnSp":
            children.append(parse_cxn_sp(child, theme, warning_log, slide_index))
        elif local == "grpSp":
            children.append(parse_grp_sp(child, theme, package, slide_part, warning_log, slide_index))
        elif local == "graphicFrame":
            table = parse_graphic_frame(child, theme, package, warning_log, slide_index)
            if table is not None:
                children.append(table)
        elif local in ("nvGrpSpPr", "grpSpPr"):
            continue
        else:
            warning_log.unsupported_element(f"p:{local}", slide_index)

    return GroupShape(
        id=shape_id,
        name=name,
        rect=rect,
        children=children,
        child_offset_emu=ch_off,
        child_extent_emu=ch_ext,
        rotation=rotation,
        flip_h=flip_h,
        flip_v=flip_v,
    )
