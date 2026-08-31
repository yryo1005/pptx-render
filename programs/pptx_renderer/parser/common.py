"""パーサー全体で共有する名前空間定数・共通処理をまとめたモジュール．

図形の座標変換（`a:xfrm`），塗りつぶし（`a:solidFill`等），
枠線（`a:ln`）の解析など，複数のパーサーから共通して利用する処理を提供する．
"""

from __future__ import annotations

from lxml import etree

from pptx_renderer.ir import Fill, RGBColor, Stroke
from pptx_renderer.parser.theme import Theme, apply_color_transforms, read_alpha, resolve_scheme_color
from pptx_renderer.units import RectEMU
from pptx_renderer.warnings_log import WarningLog

MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
A14_NS = "http://schemas.microsoft.com/office/drawing/2010/main"

NS = {"a": A_NS, "p": P_NS, "r": R_NS, "m": M_NS, "a14": A14_NS}


def qn(tag: str) -> str:
    """`prefix:local`形式のタグ名をClark記法へ変換する．"""

    prefix, local = tag.split(":")
    return f"{{{NS[prefix]}}}{local}"


def a(tag: str) -> str:
    return f"{{{A_NS}}}{tag}"


def p(tag: str) -> str:
    return f"{{{P_NS}}}{tag}"


def r(tag: str) -> str:
    return f"{{{R_NS}}}{tag}"


def emu(value: str | None, default: float = 0.0) -> float:
    """属性値（文字列）をEMU値（float）へ変換する．"""

    if value is None:
        return default
    return float(value)


def parse_xfrm(xfrm_el: etree._Element | None) -> tuple[RectEMU, float, bool, bool]:
    """`a:xfrm`要素から矩形・回転角・反転フラグを取得する．

    引数:
        xfrm_el (etree._Element | None): `a:xfrm`要素．
    戻り値:
        tuple[RectEMU, float, bool, bool]: (矩形, 回転角[度], 水平反転, 垂直反転)．
            `xfrm_el`がNoneの場合は原点0サイズの矩形を返す．
    """

    if xfrm_el is None:
        return RectEMU(0.0, 0.0, 0.0, 0.0), 0.0, False, False

    off = xfrm_el.find(a("off"))
    ext = xfrm_el.find(a("ext"))
    x = emu(off.get("x")) if off is not None else 0.0
    y = emu(off.get("y")) if off is not None else 0.0
    cx = emu(ext.get("cx")) if ext is not None else 0.0
    cy = emu(ext.get("cy")) if ext is not None else 0.0

    rot_60000ths = xfrm_el.get("rot")
    rotation = float(rot_60000ths) / 60000.0 if rot_60000ths else 0.0
    flip_h = xfrm_el.get("flipH") == "1"
    flip_v = xfrm_el.get("flipV") == "1"

    return RectEMU(x, y, cx, cy), rotation, flip_h, flip_v


def resolve_color_element(color_el: etree._Element, theme: Theme) -> tuple[RGBColor, float]:
    """`a:srgbClr`または`a:schemeClr`要素から色と不透明度を解決する．

    引数:
        color_el (etree._Element): `a:srgbClr`または`a:schemeClr`要素．
        theme (Theme): テーマ情報．
    戻り値:
        tuple[RGBColor, float]: (色, 不透明度)．
    """

    local = etree.QName(color_el).localname
    if local == "srgbClr":
        base = RGBColor.from_hex(color_el.get("val"))
    elif local == "schemeClr":
        base = resolve_scheme_color(theme, color_el.get("val"))
    else:
        base = RGBColor(0, 0, 0)

    color = apply_color_transforms(base, color_el)
    alpha = read_alpha(color_el)
    return color, alpha


def parse_fill(spPr_el: etree._Element | None, theme: Theme, warning_log: WarningLog, slide_index: int | None) -> Fill:
    """`p:spPr`（図形プロパティ）から塗りつぶし情報を解析する．

    引数:
        spPr_el (etree._Element | None): `p:spPr`要素．
        theme (Theme): テーマ情報．
        warning_log (WarningLog): 警告記録先．
        slide_index (int | None): 対象スライド番号．
    戻り値:
        Fill: 解析結果．未対応の塗り種別の場合は警告を出した上で単色近似を試みる．
    """

    if spPr_el is None:
        return Fill(kind="none")

    if spPr_el.find(a("noFill")) is not None:
        return Fill(kind="none")

    solid = spPr_el.find(a("solidFill"))
    if solid is not None:
        color_el = solid.find(a("srgbClr"))
        if color_el is None:
            color_el = solid.find(a("schemeClr"))
        if color_el is not None:
            color, alpha = resolve_color_element(color_el, theme)
            return Fill(kind="solid", color=color, alpha=alpha)

    grad = spPr_el.find(a("gradFill"))
    if grad is not None:
        warning_log.add("unsupported_fill", "グラデーション塗りを先頭の色による単色塗りへ近似しました．", slide_index)
        gs_lst = grad.find(f"{a('gsLst')}/{a('gs')}")
        if gs_lst is not None:
            color_el = gs_lst.find(a("srgbClr"))
            if color_el is None:
                color_el = gs_lst.find(a("schemeClr"))
            if color_el is not None:
                color, alpha = resolve_color_element(color_el, theme)
                return Fill(kind="solid", color=color, alpha=alpha)
        return Fill(kind="none")

    if spPr_el.find(a("blipFill")) is not None:
        warning_log.add("unsupported_fill", "画像塗り（blipFill）は未対応のため塗りなしとして扱いました．", slide_index)
        return Fill(kind="none")

    if spPr_el.find(a("pattFill")) is not None:
        warning_log.add("unsupported_fill", "パターン塗り（pattFill）は未対応のため塗りなしとして扱いました．", slide_index)
        return Fill(kind="none")

    return Fill(kind="none")


_DASH_MAP = {
    "solid": "solid",
    "dash": "dash",
    "dashDot": "dashDot",
    "lgDash": "dash",
    "lgDashDot": "dashDot",
    "lgDashDotDot": "dashDot",
    "sysDash": "dash",
    "sysDot": "dot",
    "sysDashDot": "dashDot",
    "sysDashDotDot": "dashDot",
    "dot": "dot",
}


_ARROW_TYPE_MAP = {
    "triangle": "triangle",
    "stealth": "triangle",
    "arrow": "triangle",
    "diamond": "diamond",
    "oval": "oval",
    "none": "none",
}


def _line_ref_from_style(style_el: etree._Element | None) -> tuple[int, object | None]:
    """`p:style/a:lnRef`から，参照インデックスと色要素を取得する．"""

    if style_el is None:
        return 0, None
    ln_ref = style_el.find(a("lnRef"))
    if ln_ref is None:
        return 0, None
    idx = int(ln_ref.get("idx", "0"))
    color_el = ln_ref.find(a("schemeClr"))
    if color_el is None:
        color_el = ln_ref.find(a("srgbClr"))
    return idx, color_el


def parse_line(
    spPr_el: etree._Element | None,
    theme: Theme,
    warning_log: WarningLog,
    slide_index: int | None,
    style_el: etree._Element | None = None,
) -> Stroke:
    """`p:spPr`から枠線情報を解析する．

    自身の`a:ln`が色・線幅を明示しない場合，`p:style/a:lnRef`（スタイル
    マトリクス参照）とテーマの`a:lnStyleLst`から既定値を補う．

    引数:
        spPr_el (etree._Element | None): `p:spPr`要素．
        theme (Theme): テーマ情報．
        warning_log (WarningLog): 警告記録先．
        slide_index (int | None): 対象スライド番号．
        style_el (etree._Element | None): `p:style`要素（スタイル参照の解決用）．
    戻り値:
        Stroke: 解析結果．
    """

    ln_ref_idx, ln_ref_color_el = _line_ref_from_style(style_el)
    ln_el = spPr_el.find(a("ln")) if spPr_el is not None else None

    if ln_el is None:
        if ln_ref_idx > 0 and ln_ref_color_el is not None:
            color, alpha = resolve_color_element(ln_ref_color_el, theme)
            width_pt = _line_style_width(theme, ln_ref_idx)
            return Stroke(kind="solid", color=color, width_pt=width_pt, alpha=alpha)
        return Stroke(kind="none")

    if ln_el.find(a("noFill")) is not None:
        return Stroke(kind="none")

    solid = ln_el.find(a("solidFill"))
    color_el = None
    if solid is not None:
        color_el = solid.find(a("srgbClr"))
        if color_el is None:
            color_el = solid.find(a("schemeClr"))
    if color_el is None:
        color_el = ln_ref_color_el
    if color_el is None:
        return Stroke(kind="none")

    color, alpha = resolve_color_element(color_el, theme)
    width_attr = ln_el.get("w")
    width_pt = emu(width_attr) / 12700.0 if width_attr is not None else _line_style_width(theme, ln_ref_idx)

    dash_el = ln_el.find(f"{a('prstDash')}")
    dash = _DASH_MAP.get(dash_el.get("val"), "solid") if dash_el is not None else "solid"

    head_end = ln_el.find(a("headEnd"))
    tail_end = ln_el.find(a("tailEnd"))
    head_arrow = _ARROW_TYPE_MAP.get(head_end.get("type", "none"), "none") if head_end is not None else "none"
    tail_arrow = _ARROW_TYPE_MAP.get(tail_end.get("type", "none"), "none") if tail_end is not None else "none"

    return Stroke(
        kind="solid",
        color=color,
        width_pt=width_pt,
        alpha=alpha,
        dash=dash,
        head_arrow=head_arrow,
        tail_arrow=tail_arrow,
    )


def _line_style_width(theme: Theme, ln_ref_idx: int) -> float:
    if ln_ref_idx <= 0:
        return 0.75
    widths = theme.line_style_widths_pt
    return widths[min(ln_ref_idx - 1, len(widths) - 1)]


def iter_effective_children(container_el: etree._Element):
    """`p:spTree`等の子要素を，`mc:AlternateContent`を展開しながら列挙する．

    `mc:AlternateContent`は，新しい機能（SVG画像等）を使う`mc:Choice`と，
    それを解釈できない古い実装向けの`mc:Fallback`を両方保持するラッパー要素．
    本レンダラーは`mc:Choice`側の機能に対応しているため，`mc:Choice`の
    子要素を採用し，存在しない場合のみ`mc:Fallback`の子要素を採用する．

    引数:
        container_el (etree._Element): `p:spTree`や`p:grpSp`等の親要素．
    戻り値:
        list[etree._Element]: 実体化された子要素のリスト．
    """

    result = []
    for child in container_el:
        if child.tag == f"{{{MC_NS}}}AlternateContent":
            choice = child.find(f"{{{MC_NS}}}Choice")
            target = choice if choice is not None else child.find(f"{{{MC_NS}}}Fallback")
            if target is not None:
                result.extend(list(target))
        else:
            result.append(child)
    return result


def has_cjk(text: str) -> bool:
    """文字列にCJK（中日韓）文字が含まれるかを判定する．

    引数:
        text (str): 判定対象の文字列．
    戻り値:
        bool: CJK文字が1文字でも含まれればTrue．
    """

    for ch in text:
        code = ord(ch)
        if (
            0x3040 <= code <= 0x30FF  # ひらがな・カタカナ
            or 0x3400 <= code <= 0x4DBF  # CJK拡張A
            or 0x4E00 <= code <= 0x9FFF  # CJK統合漢字
            or 0xF900 <= code <= 0xFAFF  # CJK互換漢字
            or 0xFF00 <= code <= 0xFFEF  # 全角記号
        ):
            return True
    return False
