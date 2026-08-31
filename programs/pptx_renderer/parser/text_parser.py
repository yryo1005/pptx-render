"""`p:txBody`（テキスト本体）をSlide IRの`TextBody`へ変換するパーサー．

段落（`a:p`），ラン（`a:r`），改行（`a:br`），インライン数式（`a14:m`）を
扱う．フォント名の解決はテーマのフォントパターン（`+mj-lt`等）を考慮する．
"""

from __future__ import annotations

from typing import Callable

from lxml import etree

from pptx_renderer.ir import HAlign, LineBreak, MathRun, Paragraph, RGBColor, RunIR, TextBody, TextRun, VAlign
from pptx_renderer.math.omml_to_latex import convert_omath_to_latex
from pptx_renderer.parser.common import a, resolve_color_element
from pptx_renderer.parser.placeholder import level_defaults_from_own_lst_style
from pptx_renderer.parser.theme import Theme
from pptx_renderer.warnings_log import WarningLog

_ANCHOR_MAP = {"t": VAlign.TOP, "ctr": VAlign.MIDDLE, "b": VAlign.BOTTOM}
_ALIGN_MAP = {"l": HAlign.LEFT, "ctr": HAlign.CENTER, "r": HAlign.RIGHT, "just": HAlign.JUSTIFY, "dist": HAlign.JUSTIFY}

_DEFAULT_INSET_LR = 91440.0
_DEFAULT_INSET_TB = 45720.0

# 東アジア用フォントがPPTX側で指定されていない場合に用いる既定の日本語フォント．
# 指定されたラテン用フォント（例: Calibri）が日本語グリフを持たない場合に，
# 日本語文字が無音になって消失することを防ぐためのフォールバックである．
_DEFAULT_JAPANESE_FONT = "BIZ UDGothic"


def parse_text_body(
    tx_body_el: etree._Element | None,
    theme: Theme,
    warning_log: WarningLog,
    slide_index: int | None,
    default_size_pt: float = 18.0,
    default_align: HAlign = HAlign.LEFT,
    default_color: RGBColor | None = None,
    default_bold: bool = False,
    level_defaults_fn: Callable[[int], dict] | None = None,
) -> TextBody | None:
    """`p:txBody`（または`a:txBody`）要素を`TextBody`へ変換する．

    引数:
        tx_body_el (etree._Element | None): テキスト本体要素．
        theme (Theme): テーマ情報．
        warning_log (WarningLog): 警告記録先．
        slide_index (int | None): 対象スライド番号．
        default_size_pt (float): ランにサイズ指定が無い場合の既定フォントサイズ（pt）．
        default_align (HAlign): 段落に配置指定が無い場合の既定値．
        default_color (RGBColor | None): ランに色指定が無い場合の既定色．
        default_bold (bool): ランに太字指定が無い場合の既定値．
        level_defaults_fn (Callable[[int], dict] | None): 箇条書きレベル（0始まり）を
            受け取り，スライドレイアウト／マスターから継承した既定の段落・文字書式
            （"marL_emu", "indent_emu", "align", "line_spacing_pct", "line_spacing_pt",
            "space_before_pt", "space_after_pt", "color", "bold", "size_pt"）を返す
            関数．プレースホルダでない図形の場合はNoneでよい．
    戻り値:
        TextBody | None: 変換結果．`tx_body_el`がNoneの場合はNone．
    """

    if tx_body_el is None:
        return None

    body_pr = tx_body_el.find(a("bodyPr"))
    anchor = VAlign.TOP
    wrap = True
    inset_left, inset_top, inset_right, inset_bottom = (
        _DEFAULT_INSET_LR,
        _DEFAULT_INSET_TB,
        _DEFAULT_INSET_LR,
        _DEFAULT_INSET_TB,
    )
    font_scale = 1.0

    if body_pr is not None:
        anchor = _ANCHOR_MAP.get(body_pr.get("anchor", "t"), VAlign.TOP)
        wrap = body_pr.get("wrap", "square") != "none"
        inset_left = float(body_pr.get("lIns", inset_left))
        inset_top = float(body_pr.get("tIns", inset_top))
        inset_right = float(body_pr.get("rIns", inset_right))
        inset_bottom = float(body_pr.get("bIns", inset_bottom))

        norm_autofit = body_pr.find(a("normAutofit"))
        if norm_autofit is not None and norm_autofit.get("fontScale"):
            font_scale = int(norm_autofit.get("fontScale")) / 100000.0

    # 図形自身のlstStyle（`txBody`直下）は，レイアウト／マスターのプレース
    # ホルダより優先度が高いため，両者をマージした関数を各段落へ渡す．
    own_lst_style_el = tx_body_el.find(a("lstStyle"))

    def combined_level_defaults_fn(level: int) -> dict:
        inherited = level_defaults_fn(level) if level_defaults_fn else {}
        own = level_defaults_from_own_lst_style(own_lst_style_el, level, theme)
        return {**inherited, **own}

    paragraphs = [
        _parse_paragraph(
            p_el,
            theme,
            warning_log,
            slide_index,
            default_size_pt,
            default_align,
            font_scale,
            default_color,
            default_bold,
            combined_level_defaults_fn,
        )
        for p_el in tx_body_el.findall(a("p"))
    ]

    _assign_auto_numbers(paragraphs)

    return TextBody(
        paragraphs=paragraphs,
        anchor=anchor,
        wrap=wrap,
        inset_left_emu=inset_left,
        inset_top_emu=inset_top,
        inset_right_emu=inset_right,
        inset_bottom_emu=inset_bottom,
        font_scale=font_scale,
    )


_ALIGN_NAME_TO_HALIGN = {"left": HAlign.LEFT, "center": HAlign.CENTER, "right": HAlign.RIGHT, "justify": HAlign.JUSTIFY}


def _to_roman(n: int, upper: bool) -> str:
    values = [(1000, "M"), (900, "CM"), (500, "D"), (400, "CD"), (100, "C"), (90, "XC"),
              (50, "L"), (40, "XL"), (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I")]
    result = []
    remaining = n
    for value, symbol in values:
        count, remaining = divmod(remaining, value)
        result.append(symbol * count)
    text = "".join(result)
    return text if upper else text.lower()


def _to_alpha(n: int, upper: bool) -> str:
    # 1->a, 26->z, 27->aa という表計算ソフト的な採番方式を用いる．
    letters = []
    remaining = n
    while remaining > 0:
        remaining, rem = divmod(remaining - 1, 26)
        letters.append(chr(ord("a") + rem))
    text = "".join(reversed(letters)) or "a"
    return text.upper() if upper else text


def _format_auto_number(n: int, num_fmt: str) -> str:
    """`a:buAutoNum/@type`の書式に従って番号を文字列化する．"""

    if num_fmt.startswith("arabic"):
        body = str(n)
    elif num_fmt.startswith("alphaLc"):
        body = _to_alpha(n, upper=False)
    elif num_fmt.startswith("alphaUc"):
        body = _to_alpha(n, upper=True)
    elif num_fmt.startswith("romanLc"):
        body = _to_roman(n, upper=False)
    elif num_fmt.startswith("romanUc"):
        body = _to_roman(n, upper=True)
    else:
        body = str(n)

    if "ParenR" in num_fmt:
        return f"{body})"
    if "ParenBoth" in num_fmt:
        return f"({body})"
    if "Plain" in num_fmt:
        return body
    return f"{body}."


def _assign_auto_numbers(paragraphs: list[Paragraph]) -> None:
    """`bullet_char == "#AUTO#"`の段落へ，レベルごとの連番を割り当てる．

    引数:
        paragraphs (list[Paragraph]): 対象段落のリスト（インプレースで更新する）．
    戻り値:
        なし．
    """

    counters: dict[int, int] = {}
    for paragraph in paragraphs:
        for level in list(counters.keys()):
            if level > paragraph.level:
                del counters[level]

        if paragraph.bullet_char != "#AUTO#":
            continue

        counters[paragraph.level] = counters.get(paragraph.level, 0) + 1
        paragraph.bullet_char = _format_auto_number(counters[paragraph.level], paragraph.bullet_auto_num_fmt or "arabicPeriod")


def _parse_paragraph(
    p_el: etree._Element,
    theme: Theme,
    warning_log: WarningLog,
    slide_index: int | None,
    default_size_pt: float,
    default_align: HAlign,
    font_scale: float,
    default_color: RGBColor | None,
    default_bold: bool,
    level_defaults_fn: Callable[[int], dict] | None,
) -> Paragraph:
    p_pr = p_el.find(a("pPr"))
    level = int(p_pr.get("lvl", "0")) if p_pr is not None else 0
    defaults = level_defaults_fn(level) if level_defaults_fn else {}

    if p_pr is not None and p_pr.get("algn") is not None:
        align = _ALIGN_MAP.get(p_pr.get("algn"), default_align)
    elif "align" in defaults:
        align = _ALIGN_NAME_TO_HALIGN[defaults["align"]]
    else:
        align = default_align

    marL_emu = 0.0
    if p_pr is not None and p_pr.get("marL") is not None:
        marL_emu = float(p_pr.get("marL"))
    else:
        marL_emu = defaults.get("marL_emu", 0.0)

    indent_emu = 0.0
    if p_pr is not None and p_pr.get("indent") is not None:
        indent_emu = float(p_pr.get("indent"))
    else:
        indent_emu = defaults.get("indent_emu", 0.0)

    bullet_char: str | None = None
    bullet_font: str | None = None
    bullet_auto_num_fmt: str | None = None
    if p_pr is not None and p_pr.find(a("buNone")) is not None:
        pass
    elif p_pr is not None and p_pr.find(a("buChar")) is not None:
        bu_char_el = p_pr.find(a("buChar"))
        bu_font_el = p_pr.find(a("buFont"))
        bullet_char = bu_char_el.get("char")
        bullet_font = bu_font_el.get("typeface") if bu_font_el is not None else None
    elif p_pr is not None and p_pr.find(a("buAutoNum")) is not None:
        bu_auto_num_el = p_pr.find(a("buAutoNum"))
        bullet_char = "#AUTO#"
        bullet_auto_num_fmt = bu_auto_num_el.get("type", "arabicPeriod")
    elif "bullet" in defaults and defaults["bullet"] is not None:
        bullet_char, bullet_extra = defaults["bullet"]
        if bullet_char == "#AUTO#":
            bullet_auto_num_fmt = bullet_extra
        else:
            bullet_font = bullet_extra

    line_spacing_pct: float | None = None
    line_spacing_pt: float | None = None
    ln_spc = p_pr.find(a("lnSpc")) if p_pr is not None else None
    if ln_spc is not None:
        pct_el = ln_spc.find(a("spcPct"))
        pts_el = ln_spc.find(a("spcPts"))
        if pct_el is not None:
            line_spacing_pct = int(pct_el.get("val")) / 1000.0
        elif pts_el is not None:
            line_spacing_pt = int(pts_el.get("val")) / 100.0
    else:
        line_spacing_pct = defaults.get("line_spacing_pct")
        line_spacing_pt = defaults.get("line_spacing_pt")

    spc_bef_el = p_pr.find(a("spcBef")) if p_pr is not None else None
    space_before_pt = _read_spacing(spc_bef_el) if spc_bef_el is not None else defaults.get("space_before_pt", 0.0)
    spc_aft_el = p_pr.find(a("spcAft")) if p_pr is not None else None
    space_after_pt = _read_spacing(spc_aft_el) if spc_aft_el is not None else defaults.get("space_after_pt", 0.0)

    run_default_color = defaults.get("color", default_color)
    run_default_bold = defaults.get("bold", default_bold)
    # ランに`sz`が無い場合の既定サイズは，OOXMLの書式階層（段落→lstStyle→
    # マスター）にのみ従う．「同一段落内で最初に見つかった明示的sz」を
    # 流用する実装は，スペースのみの飾りラン等が偶発的に異なるサイズを
    # 持つ場合に，本来の既定サイズを誤って上書きしてしまうため使用しない．
    default_run_size = defaults.get("size_pt", default_size_pt)

    runs: list[RunIR] = []
    for child in p_el:
        local = etree.QName(child).localname
        if local == "r":
            runs.append(_parse_run(child, theme, default_run_size, font_scale, run_default_color, run_default_bold))
        elif local == "fld":
            runs.append(_parse_run(child, theme, default_run_size, font_scale, run_default_color, run_default_bold))
        elif local == "br":
            runs.append(LineBreak())
        elif local == "m" and child.tag.startswith("{http://schemas.microsoft.com/office/drawing"):
            runs.extend(_parse_inline_math(child, warning_log, slide_index, default_run_size, font_scale))

    algn_specified = p_pr is not None and p_pr.get("algn") is not None
    if not algn_specified and runs and all(isinstance(run, MathRun) and run.display for run in runs):
        align = HAlign.CENTER

    empty_line_size_pt = None
    if not runs:
        empty_line_size_pt = default_run_size
        end_para = p_el.find(a("endParaRPr"))
        if end_para is not None and end_para.get("sz"):
            empty_line_size_pt = int(end_para.get("sz")) / 100.0 * font_scale

    return Paragraph(
        runs=runs,
        align=align,
        line_spacing_pct=line_spacing_pct,
        line_spacing_pt=line_spacing_pt,
        space_before_pt=space_before_pt,
        space_after_pt=space_after_pt,
        level=level,
        indent_left_emu=marL_emu,
        empty_line_size_pt=empty_line_size_pt,
        bullet_char=bullet_char,
        bullet_font=bullet_font,
        bullet_offset_emu=indent_emu,
        bullet_auto_num_fmt=bullet_auto_num_fmt,
    )


def _read_spacing(spc_el: etree._Element | None) -> float:
    if spc_el is None:
        return 0.0
    pts_el = spc_el.find(a("spcPts"))
    if pts_el is not None:
        return int(pts_el.get("val")) / 100.0
    return 0.0


def _resolve_font_names(r_pr: etree._Element | None, theme: Theme) -> tuple[str, str]:
    """ランのラテン用・東アジア用フォント名を，文字ごとの選択に使えるよう個別に解決する．

    引数:
        r_pr (etree._Element | None): `a:rPr`要素．
        theme (Theme): テーマ情報．
    戻り値:
        tuple[str, str]: (ラテン用フォント名, 東アジア用フォント名)．
            東アジア用が未指定・不明な場合は既定の日本語フォントへフォールバックする．
    """

    latin_face = None
    ea_face = None
    if r_pr is not None:
        latin_el = r_pr.find(a("latin"))
        ea_el = r_pr.find(a("ea"))
        if latin_el is not None:
            latin_face = latin_el.get("typeface")
        if ea_el is not None:
            ea_face = ea_el.get("typeface")

    def resolve_theme_ref(face: str | None) -> str | None:
        if face is None:
            return None
        if face == "+mj-lt":
            return theme.font_scheme.major_latin
        if face == "+mn-lt":
            return theme.font_scheme.minor_latin
        if face == "+mj-ea":
            return theme.font_scheme.major_ea or theme.font_scheme.major_latin
        if face == "+mn-ea":
            return theme.font_scheme.minor_ea or theme.font_scheme.minor_latin
        return face

    latin_resolved = resolve_theme_ref(latin_face) or theme.font_scheme.minor_latin or "Calibri"
    ea_resolved = resolve_theme_ref(ea_face) or theme.font_scheme.minor_ea or _DEFAULT_JAPANESE_FONT

    return latin_resolved, ea_resolved


def _parse_run(
    r_el: etree._Element,
    theme: Theme,
    default_size_pt: float,
    font_scale: float,
    default_color: RGBColor | None = None,
    default_bold: bool = False,
) -> TextRun:
    r_pr = r_el.find(a("rPr"))
    t_el = r_el.find(a("t"))
    text = t_el.text if t_el is not None and t_el.text is not None else ""

    size_pt = default_size_pt
    bold = default_bold
    italic = False
    underline = False
    color = default_color if default_color is not None else theme.color_scheme.get("dk1", RGBColor(0, 0, 0))
    alpha = 1.0
    char_spacing_pt = 0.0

    if r_pr is not None:
        if r_pr.get("sz"):
            size_pt = int(r_pr.get("sz")) / 100.0
        if r_pr.get("spc") is not None:
            char_spacing_pt = int(r_pr.get("spc")) / 100.0
        bold = r_pr.get("b") == "1"
        italic = r_pr.get("i") == "1"
        underline = r_pr.get("u") not in (None, "none")

        solid_fill = r_pr.find(a("solidFill"))
        if solid_fill is not None:
            color_el = solid_fill.find(a("srgbClr"))
            if color_el is None:
                color_el = solid_fill.find(a("schemeClr"))
            if color_el is not None:
                color, alpha = resolve_color_element(color_el, theme)

    latin_font, ea_font = _resolve_font_names(r_pr, theme)

    return TextRun(
        text=text,
        font_name=latin_font,
        ea_font_name=ea_font,
        size_pt=size_pt * font_scale,
        char_spacing_pt=char_spacing_pt * font_scale,
        bold=bold,
        italic=italic,
        underline=underline,
        color=color,
        alpha=alpha,
    )


_M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"


def _m(tag: str) -> str:
    return f"{{{_M_NS}}}{tag}"


def _omath_size_pt(omath_el: etree._Element, default_size_pt: float) -> float:
    for el in omath_el.iter():
        if etree.QName(el).localname == "rPr" and el.tag.startswith("{http://schemas.openxmlformats.org/drawingml"):
            if el.get("sz"):
                return int(el.get("sz")) / 100.0
    return default_size_pt


def _parse_inline_math(
    a14_m_el: etree._Element,
    warning_log: WarningLog,
    slide_index: int | None,
    default_size_pt: float,
    font_scale: float,
) -> list[RunIR]:
    oMathPara = a14_m_el.find(_m("oMathPara"))
    if oMathPara is not None:
        result: list[RunIR] = []
        for omath in oMathPara.findall(_m("oMath")):
            size_pt = _omath_size_pt(omath, default_size_pt) * font_scale
            latex = convert_omath_to_latex(omath, warning_log, slide_index)
            result.append(MathRun(latex_body=latex, size_pt=size_pt, display=True))
        return result

    omath = a14_m_el.find(_m("oMath"))
    if omath is not None:
        size_pt = _omath_size_pt(omath, default_size_pt) * font_scale
        latex = convert_omath_to_latex(omath, warning_log, slide_index)
        return [MathRun(latex_body=latex, size_pt=size_pt, display=False)]

    return []
