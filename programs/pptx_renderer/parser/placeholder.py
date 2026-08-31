"""プレースホルダ図形の位置・サイズ・既定書式をスライドレイアウト／マスターから
継承解決するモジュール．

PowerPointでは，プレースホルダの位置・サイズ・段落書式（インデント・行間・
段落前後の空白）・文字書式（色・太字・サイズ）を，スライド自身が明示的に
上書きしない限り，以下の優先順位で継承する．

```text
1. スライド自身の指定（本モジュールの対象外，呼び出し側で処理）
2. スライドレイアウト側の，同じプレースホルダ（idx一致，無ければtype一致）のlstStyle
3. スライドマスター側の，同じプレースホルダのlstStyle
4. スライドマスターの txStyles（titleStyle/bodyStyle/otherStyle）の対応レベル
```

この継承チェーンを解決しない場合，インデント・段落前後の空白（`spcBef`等）が
既定値（0）のまま扱われ，実際のPowerPoint表示より詰まった見た目になる．
"""

from __future__ import annotations

from lxml import etree

from pptx_renderer.ir import RGBColor
from pptx_renderer.parser.common import a, p, resolve_color_element
from pptx_renderer.parser.package import PptxPackage
from pptx_renderer.parser.theme import Theme
from pptx_renderer.units import RectEMU


def _find_placeholder_sp(sp_tree: etree._Element, ph_type: str | None, ph_idx: str | None) -> etree._Element | None:
    candidates = sp_tree.findall(p("sp"))

    if ph_idx is not None:
        for sp in candidates:
            ph = sp.find(f"{p('nvSpPr')}/{p('nvPr')}/{p('ph')}")
            if ph is not None and ph.get("idx") == ph_idx:
                return sp

    if ph_type is not None:
        for sp in candidates:
            ph = sp.find(f"{p('nvSpPr')}/{p('nvPr')}/{p('ph')}")
            if ph is not None and ph.get("type", "body") == ph_type:
                return sp

    return None


def resolve_placeholder_rect(
    package: PptxPackage, slide_part: str, ph_type: str | None, ph_idx: str | None
) -> RectEMU | None:
    """レイアウト→マスターの順に，一致するプレースホルダの矩形を探索する．

    引数:
        package (PptxPackage): PPTXパッケージ．
        slide_part (str): 対象スライドのパート名．
        ph_type (str | None): `p:ph@type`（例: "title", "body"）．
        ph_idx (str | None): `p:ph@idx`．
    戻り値:
        RectEMU | None: 解決できた矩形．解決できない場合はNone．
    """

    layout_part = package.slide_layout_for_slide(slide_part)
    if layout_part is not None:
        layout_root = package.read_xml(layout_part)
        sp_tree = layout_root.find(f"{p('cSld')}/{p('spTree')}")
        if sp_tree is not None:
            sp = _find_placeholder_sp(sp_tree, ph_type, ph_idx)
            if sp is not None:
                xfrm = sp.find(f"{p('spPr')}/{a('xfrm')}")
                rect = _rect_from_xfrm(xfrm)
                if rect is not None:
                    return rect

        master_part = package.slide_master_for_layout(layout_part)
        if master_part is not None:
            master_root = package.read_xml(master_part)
            sp_tree = master_root.find(f"{p('cSld')}/{p('spTree')}")
            if sp_tree is not None:
                sp = _find_placeholder_sp(sp_tree, ph_type, ph_idx)
                if sp is not None:
                    xfrm = sp.find(f"{p('spPr')}/{a('xfrm')}")
                    rect = _rect_from_xfrm(xfrm)
                    if rect is not None:
                        return rect

    return None


_TITLE_TYPES = {"title", "ctrTitle"}
_BODY_TYPES = {"body", "subTitle", None}

# ランレベルの書式（defRPr）から取得するキー．
_RUN_KEYS = ("color", "bold", "size_pt")
# 段落レベルの書式（lvlNpPr自身の属性）から取得するキー．
_PARA_KEYS = (
    "marL_emu", "indent_emu", "align", "line_spacing_pct", "line_spacing_pt",
    "space_before_pt", "space_after_pt", "bullet",
)

_ALIGN_MAP = {"l": "left", "ctr": "center", "r": "right", "just": "justify", "dist": "justify"}


def _extract_rpr_props(rpr: etree._Element | None, theme: Theme) -> dict:
    if rpr is None:
        return {}
    props = {}
    if rpr.get("b") is not None:
        props["bold"] = rpr.get("b") == "1"
    if rpr.get("sz") is not None:
        props["size_pt"] = int(rpr.get("sz")) / 100.0
    solid = rpr.find(a("solidFill"))
    if solid is not None:
        color_el = solid.find(a("srgbClr"))
        if color_el is None:
            color_el = solid.find(a("schemeClr"))
        if color_el is not None:
            color, _ = resolve_color_element(color_el, theme)
            props["color"] = color
    return props


def _extract_ppr_props(lvl_ppr: etree._Element | None) -> dict:
    """`a:lvlNpPr`要素自身が持つ段落レベルの書式（属性・spcBef等）を抽出する．"""

    if lvl_ppr is None:
        return {}
    props: dict = {}
    if lvl_ppr.get("marL") is not None:
        props["marL_emu"] = float(lvl_ppr.get("marL"))
    if lvl_ppr.get("indent") is not None:
        props["indent_emu"] = float(lvl_ppr.get("indent"))
    if lvl_ppr.get("algn") is not None:
        props["align"] = _ALIGN_MAP.get(lvl_ppr.get("algn"), "left")

    ln_spc = lvl_ppr.find(a("lnSpc"))
    if ln_spc is not None:
        pct_el = ln_spc.find(a("spcPct"))
        pts_el = ln_spc.find(a("spcPts"))
        if pct_el is not None:
            props["line_spacing_pct"] = int(pct_el.get("val")) / 1000.0
        elif pts_el is not None:
            props["line_spacing_pt"] = int(pts_el.get("val")) / 100.0

    spc_bef = lvl_ppr.find(a("spcBef"))
    if spc_bef is not None:
        pts_el = spc_bef.find(a("spcPts"))
        if pts_el is not None:
            props["space_before_pt"] = int(pts_el.get("val")) / 100.0

    spc_aft = lvl_ppr.find(a("spcAft"))
    if spc_aft is not None:
        pts_el = spc_aft.find(a("spcPts"))
        if pts_el is not None:
            props["space_after_pt"] = int(pts_el.get("val")) / 100.0

    if lvl_ppr.find(a("buNone")) is not None:
        props["bullet"] = None
    else:
        bu_char_el = lvl_ppr.find(a("buChar"))
        bu_auto_num_el = lvl_ppr.find(a("buAutoNum"))
        if bu_char_el is not None:
            bu_font_el = lvl_ppr.find(a("buFont"))
            props["bullet"] = (bu_char_el.get("char"), bu_font_el.get("typeface") if bu_font_el is not None else None)
        elif bu_auto_num_el is not None:
            props["bullet"] = ("#AUTO#", bu_auto_num_el.get("type", "arabicPeriod"))

    return props


def _level_props_from_placeholder_sp(sp: etree._Element, level: int, theme: Theme) -> dict:
    tx_body = sp.find(f"{p('txBody')}")
    if tx_body is None:
        return {}

    lvl_tag = f"lvl{level + 1}pPr"
    lvl_el = tx_body.find(f"{a('lstStyle')}/{a(lvl_tag)}")
    props = _extract_ppr_props(lvl_el)
    props.update(_extract_rpr_props(lvl_el.find(a("defRPr")) if lvl_el is not None else None, theme))

    if level == 0:
        # レベル0（既定段落）については，プレースホルダ本体の最初の段落に
        # 書かれた endParaRPr / pPr/defRPr が実質的な既定値になっている場合がある．
        first_p = tx_body.find(a("p"))
        if first_p is not None:
            p_pr = first_p.find(a("pPr"))
            if p_pr is not None:
                merged = _extract_rpr_props(p_pr.find(a("defRPr")), theme)
                props = {**merged, **props}
            end_para = first_p.find(a("endParaRPr"))
            merged = _extract_rpr_props(end_para, theme)
            props = {**merged, **props}

    return props


def level_defaults_from_own_lst_style(lst_style_el: etree._Element | None, level: int, theme: Theme) -> dict:
    """図形自身の`a:lstStyle`（`txBody`直下）から，指定レベルの既定書式を解決する．

    OOXMLの書式解決順序では，図形自身の`a:lstStyle`は，スライドレイアウト／
    マスターのプレースホルダより優先度が高い．`sample.pptx`では，本文の
    フォントサイズ（`sz`）が図形自身の`lstStyle`にのみ明示され，レイアウト／
    マスター側には無いケースが確認されており，本関数を経由しないと既定
    サイズを取得できない．

    引数:
        lst_style_el (etree._Element | None): 図形の`p:txBody/a:lstStyle`要素．
        level (int): 段落の箇条書きレベル（0始まり）．
        theme (Theme): テーマ情報（色解決用）．
    戻り値:
        dict: `resolve_placeholder_level_defaults`と同じ形式の辞書．
            `lst_style_el`がNone，または対応レベルの定義が無い場合は空辞書．
    """

    if lst_style_el is None:
        return {}
    lvl_el = lst_style_el.find(a(f"lvl{level + 1}pPr"))
    if lvl_el is None:
        return {}
    props = _extract_ppr_props(lvl_el)
    props.update(_extract_rpr_props(lvl_el.find(a("defRPr")), theme))
    return props


def resolve_placeholder_level_defaults(
    package: PptxPackage,
    slide_part: str,
    ph_type: str | None,
    ph_idx: str | None,
    level: int,
    theme: Theme,
) -> dict:
    """プレースホルダの既定書式（文字・段落）を，指定レベルについて解決する．

    優先順位は「スライドレイアウトの同一プレースホルダ」→「スライドマスターの
    同一プレースホルダ」→「スライドマスターの txStyles（対応レベル）」である．

    引数:
        package (PptxPackage): PPTXパッケージ．
        slide_part (str): 対象スライドのパート名．
        ph_type (str | None): `p:ph@type`．
        ph_idx (str | None): `p:ph@idx`．
        level (int): 段落の箇条書きレベル（0始まり）．
        theme (Theme): テーマ情報（色解決用）．
    戻り値:
        dict: "color", "bold", "size_pt", "marL_emu", "indent_emu", "align",
            "line_spacing_pct", "line_spacing_pt", "space_before_pt",
            "space_after_pt" のうち，解決できたキーのみを含む辞書．
            より優先度の高い階層で見つかった値が優先される．
    """

    all_keys = set(_RUN_KEYS) | set(_PARA_KEYS)
    result: dict = {}

    layout_part = package.slide_layout_for_slide(slide_part)
    if layout_part is None:
        return result

    layout_root = package.read_xml(layout_part)
    layout_sp_tree = layout_root.find(f"{p('cSld')}/{p('spTree')}")
    if layout_sp_tree is not None:
        sp = _find_placeholder_sp(layout_sp_tree, ph_type, ph_idx)
        if sp is not None:
            result = _level_props_from_placeholder_sp(sp, level, theme)

    if all_keys.issubset(result.keys()):
        return result

    master_part = package.slide_master_for_layout(layout_part)
    if master_part is None:
        return result
    master_root = package.read_xml(master_part)

    master_sp_tree = master_root.find(f"{p('cSld')}/{p('spTree')}")
    if master_sp_tree is not None:
        sp = _find_placeholder_sp(master_sp_tree, ph_type, ph_idx)
        if sp is not None:
            fallback = _level_props_from_placeholder_sp(sp, level, theme)
            result = {**fallback, **result}

    if all_keys.issubset(result.keys()):
        return result

    style_tag = "titleStyle" if ph_type in _TITLE_TYPES else ("bodyStyle" if ph_type in _BODY_TYPES else "otherStyle")
    tx_styles = master_root.find(f"{p('txStyles')}/{p(style_tag)}")
    if tx_styles is not None:
        lvl_el = tx_styles.find(a(f"lvl{level + 1}pPr"))
        if lvl_el is not None:
            fallback = _extract_ppr_props(lvl_el)
            fallback.update(_extract_rpr_props(lvl_el.find(a("defRPr")), theme))
            result = {**fallback, **result}

    return result


def _rect_from_xfrm(xfrm: etree._Element | None) -> RectEMU | None:
    if xfrm is None:
        return None
    off = xfrm.find(a("off"))
    ext = xfrm.find(a("ext"))
    if off is None or ext is None:
        return None
    return RectEMU(
        x=float(off.get("x")), y=float(off.get("y")), cx=float(ext.get("cx")), cy=float(ext.get("cy"))
    )
