"""テーマ（配色パターン・フォントパターン）を解析するモジュール．

`a:schemeClr`（配色）および`+mj-lt`/`+mn-ea`等（テーマフォント参照）を
実際の色・フォント名へ解決する処理をここに集約する．
"""

from __future__ import annotations

import colorsys
from dataclasses import dataclass, field

from lxml import etree

from pptx_renderer.ir import RGBColor

_A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"


def _a(tag: str) -> str:
    return f"{{{_A_NS}}}{tag}"


@dataclass
class FontScheme:
    """テーマのフォントパターン（見出し用・本文用）．"""

    major_latin: str = "Calibri"
    major_ea: str = ""
    minor_latin: str = "Calibri"
    minor_ea: str = ""


@dataclass
class Theme:
    """1つのテーマから抽出した配色・フォント情報．

    属性:
        color_scheme (dict[str, RGBColor]): "dk1","lt1","dk2","lt2",
            "accent1"〜"accent6","hlink","folHlink"をキーとする色の辞書．
        font_scheme (FontScheme): フォントパターン．
        line_style_widths_pt (list[float]): `a:fmtScheme/a:lnStyleLst`の各
            線幅（pt）．インデックス0が`a:lnRef idx="1"`に対応する．
            図形の`a:ln`に`w`属性が無く，`p:style/a:lnRef`で線スタイルが
            参照されている場合の既定線幅として使用する．
    """

    color_scheme: dict[str, RGBColor] = field(default_factory=dict)
    font_scheme: FontScheme = field(default_factory=FontScheme)
    line_style_widths_pt: list[float] = field(default_factory=lambda: [0.75, 1.5, 2.25])


_DEFAULT_COLOR_SCHEME = {
    "dk1": RGBColor(0, 0, 0),
    "lt1": RGBColor(255, 255, 255),
    "dk2": RGBColor(68, 68, 68),
    "lt2": RGBColor(238, 238, 238),
    "accent1": RGBColor(68, 114, 196),
    "accent2": RGBColor(237, 125, 49),
    "accent3": RGBColor(165, 165, 165),
    "accent4": RGBColor(255, 192, 0),
    "accent5": RGBColor(91, 155, 213),
    "accent6": RGBColor(112, 173, 71),
    "hlink": RGBColor(5, 99, 193),
    "folHlink": RGBColor(149, 79, 114),
}


def parse_theme(theme_root: etree._Element | None) -> Theme:
    """`theme1.xml`のルート要素からTheme情報を抽出する．

    引数:
        theme_root (etree._Element | None): テーマXMLのルート要素．Noneの場合は既定値を返す．
    戻り値:
        Theme: 抽出結果．
    """

    theme = Theme(color_scheme=dict(_DEFAULT_COLOR_SCHEME))
    if theme_root is None:
        return theme

    clr_scheme = theme_root.find(f"{_a('themeElements')}/{_a('clrScheme')}")
    if clr_scheme is not None:
        for key in list(_DEFAULT_COLOR_SCHEME.keys()):
            node = clr_scheme.find(_a(key))
            if node is None:
                continue
            color = _read_color_node(node)
            if color is not None:
                theme.color_scheme[key] = color

    font_scheme_el = theme_root.find(f"{_a('themeElements')}/{_a('fontScheme')}")
    if font_scheme_el is not None:
        major = font_scheme_el.find(f"{_a('majorFont')}")
        minor = font_scheme_el.find(f"{_a('minorFont')}")
        if major is not None:
            theme.font_scheme.major_latin = _typeface(major, "latin", theme.font_scheme.major_latin)
            theme.font_scheme.major_ea = _typeface(major, "ea", "")
        if minor is not None:
            theme.font_scheme.minor_latin = _typeface(minor, "latin", theme.font_scheme.minor_latin)
            theme.font_scheme.minor_ea = _typeface(minor, "ea", "")

    ln_style_lst = theme_root.find(f"{_a('themeElements')}/{_a('fmtScheme')}/{_a('lnStyleLst')}")
    if ln_style_lst is not None:
        widths = []
        for ln_el in ln_style_lst.findall(_a("ln")):
            w = ln_el.get("w")
            widths.append(float(w) / 12700.0 if w is not None else 0.75)
        if widths:
            theme.line_style_widths_pt = widths

    return theme


def _typeface(font_el: etree._Element, script: str, default: str) -> str:
    node = font_el.find(_a(script))
    if node is None:
        return default
    return node.get("typeface", default)


def _read_color_node(container: etree._Element) -> RGBColor | None:
    """`a:dk1`等の色コンテナ要素から，実際の色定義（`a:srgbClr`/`a:sysClr`）を読み取る．"""

    srgb = container.find(_a("srgbClr"))
    if srgb is not None:
        return RGBColor.from_hex(srgb.get("val"))
    sys_clr = container.find(_a("sysClr"))
    if sys_clr is not None:
        last_clr = sys_clr.get("lastClr")
        if last_clr:
            return RGBColor.from_hex(last_clr)
    return None


# schemeClrのval属性から，Themeのcolor_schemeキーへの既定マッピング．
# （p:clrMapによるカスタム再マップは本実装では簡略化のため未対応．）
_SCHEME_CLR_ALIAS = {
    "tx1": "dk1",
    "tx2": "dk2",
    "bg1": "lt1",
    "bg2": "lt2",
    "dk1": "dk1",
    "dk2": "dk2",
    "lt1": "lt1",
    "lt2": "lt2",
    "accent1": "accent1",
    "accent2": "accent2",
    "accent3": "accent3",
    "accent4": "accent4",
    "accent5": "accent5",
    "accent6": "accent6",
    "hlink": "hlink",
    "folHlink": "folHlink",
}


def resolve_scheme_color(theme: Theme, scheme_val: str) -> RGBColor:
    """`a:schemeClr`の`val`属性から実際の色を解決する（変換なし）．

    引数:
        theme (Theme): テーマ情報．
        scheme_val (str): `a:schemeClr@val`の値（例: "accent1", "tx1"）．
    戻り値:
        RGBColor: 解決された色．未知の値の場合は黒を返す．
    """

    key = _SCHEME_CLR_ALIAS.get(scheme_val, scheme_val)
    return theme.color_scheme.get(key, RGBColor(0, 0, 0))


def apply_color_transforms(color: RGBColor, clr_element: etree._Element) -> RGBColor:
    """`a:schemeClr`等の子要素として指定される色変換（lumMod/lumOff/shade/tint/alpha）を適用する．

    引数:
        color (RGBColor): 変換前の色．
        clr_element (etree._Element): 色を指定する要素（例: `a:schemeClr`）．
            子要素として `a:lumMod`, `a:lumOff`, `a:shade`, `a:tint` を持ちうる．
    戻り値:
        RGBColor: 変換後の色．
    """

    r, g, b = color.r / 255.0, color.g / 255.0, color.b / 255.0
    h, l, s = colorsys.rgb_to_hls(r, g, b)[0], colorsys.rgb_to_hls(r, g, b)[1], colorsys.rgb_to_hls(r, g, b)[2]

    lum_mod_el = clr_element.find(_a("lumMod"))
    lum_off_el = clr_element.find(_a("lumOff"))
    if lum_mod_el is not None or lum_off_el is not None:
        mod = int(lum_mod_el.get("val")) / 100000.0 if lum_mod_el is not None else 1.0
        off = int(lum_off_el.get("val")) / 100000.0 if lum_off_el is not None else 0.0
        l = max(0.0, min(1.0, l * mod + off))

    shade_el = clr_element.find(_a("shade"))
    if shade_el is not None:
        factor = int(shade_el.get("val")) / 100000.0
        l = l * factor

    tint_el = clr_element.find(_a("tint"))
    if tint_el is not None:
        factor = int(tint_el.get("val")) / 100000.0
        l = l * factor + (1.0 - factor)

    r2, g2, b2 = colorsys.hls_to_rgb(h, l, s)
    return RGBColor(round(r2 * 255), round(g2 * 255), round(b2 * 255))


def read_alpha(clr_element: etree._Element) -> float:
    """色要素の`a:alpha`子要素から不透明度を読み取る．

    引数:
        clr_element (etree._Element): 色を指定する要素．
    戻り値:
        float: 不透明度（0.0〜1.0）．指定が無い場合は1.0．
    """

    alpha_el = clr_element.find(_a("alpha"))
    if alpha_el is None:
        return 1.0
    return int(alpha_el.get("val")) / 100000.0
