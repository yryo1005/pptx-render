"""OMML（Office Math Markup Language）をLaTeXの数式文字列へ変換するモジュール．

PowerPointのネイティブ数式（`m:oMath`）を，独立したIR（LaTeX文字列）へ変換する．
このモジュールはLaTeX文字列の生成のみを担当し，実際のPDF/SVGへのレンダリングは
`pptx_renderer.math.latex_render` が担当する．

対応する主なOMML要素:
    m:r（数式ラン），m:sSub/m:sSup/m:sSubSup（上付き・下付き），m:f（分数），
    m:d（区切り記号），m:rad（根号），m:nary（総和・積分等のN項演算子），
    m:func（関数），m:limLow/m:limUpp（極限），m:acc（アクセント），
    m:bar（上線・下線），m:groupChr（上括弧・下括弧），m:eqArr（数式配列），
    m:m（行列），m:sPre（前置上付き・下付き），m:box/m:borderBox．

未対応の要素に遭遇した場合は，レンダリングを中断せず，警告を記録した上で
子要素のテキストのみを連結して処理を継続する．
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field

from lxml import etree

from pptx_renderer.warnings_log import WarningLog, default_log

_M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
_A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"

NSMAP = {"m": _M_NS, "a": _A_NS}


def qn(tag: str) -> str:
    """`m:xxx` または `a:xxx` 形式のタグ名をClark記法へ変換する．

    引数:
        tag (str): 名前空間プレフィックス付きのタグ名（例: "m:oMath"）．
    戻り値:
        str: Clark記法のタグ名（例: "{http://...}oMath"）．
    """

    prefix, local = tag.split(":")
    return f"{{{NSMAP[prefix]}}}{local}"


# よく使われるn項演算子記号からLaTeXコマンドへの対応表．
_NARY_CHR_MAP = {
    "∑": r"\sum",
    "∏": r"\prod",
    "∫": r"\int",
    "∬": r"\iint",
    "∭": r"\iiint",
    "∮": r"\oint",
    "⋃": r"\bigcup",
    "⋂": r"\bigcap",
    "⨁": r"\bigoplus",
    "⨂": r"\bigotimes",
    "⨄": r"\biguplus",
    "⨆": r"\bigsqcup",
    "⋀": r"\bigwedge",
    "⋁": r"\bigvee",
}

# 区切り文字（m:begChr / m:endChr）からLaTeXの \left \right 用トークンへの対応表．
_DELIM_MAP = {
    "(": "(",
    ")": ")",
    "[": "[",
    "]": "]",
    "{": r"\{",
    "}": r"\}",
    "|": "|",
    "‖": r"\|",
    "〈": r"\langle",
    "〉": r"\rangle",
    "⌊": r"\lfloor",
    "⌋": r"\rfloor",
    "⌈": r"\lceil",
    "⌉": r"\rceil",
    "": ".",
}

# アクセント文字（結合文字を含む）からLaTeXアクセントコマンドへの対応表．
_ACCENT_MAP = {
    "̀": r"\grave",
    "́": r"\acute",
    "̂": r"\hat",
    "̃": r"\tilde",
    "̄": r"\bar",
    "̅": r"\overline",
    "̇": r"\dot",
    "̈": r"\ddot",
    "⃗": r"\vec",
    "→": r"\vec",
    "ˇ": r"\check",
    "˘": r"\breve",
}

# LaTeXの組み込み演算子名として直接使用できる関数名．
_KNOWN_FUNC_NAMES = {
    "sin", "cos", "tan", "cot", "sec", "csc",
    "sinh", "cosh", "tanh", "coth",
    "arcsin", "arccos", "arctan",
    "log", "ln", "exp", "det", "gcd", "deg",
    "max", "min", "sup", "inf", "lim", "arg",
}

_MATH_SPECIAL_CHARS = str.maketrans({
    "\\": r"\textbackslash ",
    "{": r"\{",
    "}": r"\}",
    "_": r"\_",
    "^": r"\^{}",
    "%": r"\%",
    "#": r"\#",
    "&": r"\&",
    "$": r"\$",
})


def _escape_text(text: str) -> str:
    """LaTeXの数式モードで特別な意味を持つ文字をエスケープする．

    引数:
        text (str): OMML上の生テキスト．
    戻り値:
        str: LaTeXの数式モードにそのまま埋め込める文字列．
    """

    return text.translate(_MATH_SPECIAL_CHARS)


@dataclass
class OmmlConversionContext:
    """OMML変換処理中に共有する状態を保持するデータクラス．

    属性:
        warning_log (WarningLog): 未対応要素を記録する警告ロガー．
        slide_index (int | None): 変換対象のスライド番号．
        unsupported_tags (set[str]): 既に警告済みの未対応タグ名の集合（重複警告の抑制用）．
    """

    warning_log: WarningLog = field(default_factory=lambda: default_log)
    slide_index: int | None = None
    unsupported_tags: set[str] = field(default_factory=set)

    def warn_unsupported(self, tag: str) -> None:
        """未対応のOMML要素について，重複を避けつつ警告を記録する．

        引数:
            tag (str): 未対応であった要素のタグ名（例: "m:sPre"）．
        戻り値:
            なし．
        """

        key = f"{tag}@{self.slide_index}"
        if key in self.unsupported_tags:
            return
        self.unsupported_tags.add(key)
        self.warning_log.unsupported_element(f"OMML element {tag}", self.slide_index)


def _find(element: etree._Element, tag: str) -> etree._Element | None:
    return element.find(qn(tag))


def _findall_direct(element: etree._Element, tags: tuple[str, ...]) -> list[etree._Element]:
    """指定タグ集合のうち，直下の子要素のみを順序を保って取得する．"""

    wanted = {qn(t) for t in tags}
    return [child for child in element if child.tag in wanted]


# 数式コンテンツとして子要素を持ちうる要素タグの一覧．
_CONTENT_TAGS = (
    "m:r", "m:sSub", "m:sSup", "m:sSubSup", "m:sPre", "m:f", "m:d", "m:rad",
    "m:nary", "m:func", "m:limLow", "m:limUpp", "m:acc", "m:bar",
    "m:groupChr", "m:eqArr", "m:m", "m:box", "m:borderBox",
)


def _convert_run(element: etree._Element, ctx: OmmlConversionContext) -> str:
    """`m:r`（数式ラン）をLaTeXへ変換する．"""

    text_parts = [t.text or "" for t in element.findall(qn("m:t"))]
    text = "".join(text_parts)
    if text == "":
        return ""

    a_rpr = element.find(qn("a:rPr"))
    bold = False
    italic = None  # None: 明示指定なし（既定の数式斜体に任せる）
    color_hex = None
    if a_rpr is not None:
        if a_rpr.get("b") == "1":
            bold = True
        if a_rpr.get("i") is not None:
            italic = a_rpr.get("i") == "1"
        fill = a_rpr.find(qn("a:solidFill"))
        if fill is not None:
            srgb = fill.find(qn("a:srgbClr"))
            if srgb is not None:
                color_hex = srgb.get("val")

    is_ascii_letters = re.fullmatch(r"[A-Za-z]+", text) is not None
    escaped = _escape_text(text)

    if is_ascii_letters:
        if bold and italic is False:
            body = rf"\mathbf{{{escaped}}}"
        elif bold and italic is not False:
            body = rf"\boldsymbol{{{escaped}}}"
        elif not bold and italic is False:
            body = rf"\mathrm{{{escaped}}}"
        else:
            body = escaped
    else:
        body = escaped

    if color_hex:
        body = rf"\textcolor[HTML]{{{color_hex.upper()}}}{{{body}}}"

    return body


def _convert_children(element: etree._Element, ctx: OmmlConversionContext) -> str:
    """要素直下にある数式コンテンツ子要素を順に変換し，連結する．"""

    parts = []
    for child in _findall_direct(element, _CONTENT_TAGS):
        parts.append(_convert_node(child, ctx))
    return " ".join(p for p in parts if p)


def _convert_e(parent: etree._Element, ctx: OmmlConversionContext) -> str:
    """`m:e`（基底要素）を1つ取り出して変換する．"""

    e = _find(parent, "m:e")
    if e is None:
        return ""
    return _convert_children(e, ctx)


def _convert_sub_sup_text(parent: etree._Element, tag: str, ctx: OmmlConversionContext) -> str:
    node = _find(parent, tag)
    if node is None:
        return ""
    return _convert_children(node, ctx)


def _convert_sSub(element: etree._Element, ctx: OmmlConversionContext) -> str:
    base = _convert_e(element, ctx)
    sub = _convert_sub_sup_text(element, "m:sub", ctx)
    return f"{{{base}}}_{{{sub}}}"


def _convert_sSup(element: etree._Element, ctx: OmmlConversionContext) -> str:
    base = _convert_e(element, ctx)
    sup = _convert_sub_sup_text(element, "m:sup", ctx)
    return f"{{{base}}}^{{{sup}}}"


def _convert_sSubSup(element: etree._Element, ctx: OmmlConversionContext) -> str:
    base = _convert_e(element, ctx)
    sub = _convert_sub_sup_text(element, "m:sub", ctx)
    sup = _convert_sub_sup_text(element, "m:sup", ctx)
    return f"{{{base}}}_{{{sub}}}^{{{sup}}}"


def _convert_sPre(element: etree._Element, ctx: OmmlConversionContext) -> str:
    base = _convert_e(element, ctx)
    sub = _convert_sub_sup_text(element, "m:sub", ctx)
    sup = _convert_sub_sup_text(element, "m:sup", ctx)
    return f"{{}}_{{{sub}}}^{{{sup}}}{{{base}}}"


def _convert_f(element: etree._Element, ctx: OmmlConversionContext) -> str:
    f_pr = _find(element, "m:fPr")
    frac_type = "bar"
    if f_pr is not None:
        type_node = _find(f_pr, "m:type")
        if type_node is not None:
            frac_type = type_node.get("{%s}val" % _M_NS, frac_type)

    num_node = _find(element, "m:num")
    den_node = _find(element, "m:den")
    num = _convert_children(num_node, ctx) if num_node is not None else ""
    den = _convert_children(den_node, ctx) if den_node is not None else ""

    if frac_type in ("lin", "skw"):
        return f"{{{num}}}/{{{den}}}"
    return rf"\frac{{{num}}}{{{den}}}"


def _convert_d(element: etree._Element, ctx: OmmlConversionContext) -> str:
    d_pr = _find(element, "m:dPr")
    beg_chr = "("
    end_chr = ")"
    sep_chr = ","
    if d_pr is not None:
        beg_node = _find(d_pr, "m:begChr")
        end_node = _find(d_pr, "m:endChr")
        sep_node = _find(d_pr, "m:sepChr")
        if beg_node is not None:
            beg_chr = beg_node.get("{%s}val" % _M_NS, beg_chr)
        if end_node is not None:
            end_chr = end_node.get("{%s}val" % _M_NS, end_chr)
        if sep_node is not None:
            sep_chr = sep_node.get("{%s}val" % _M_NS, sep_chr)

    entries = element.findall(qn("m:e"))
    inner = f" {sep_chr} ".join(_convert_children(e, ctx) for e in entries)

    left = _DELIM_MAP.get(beg_chr, beg_chr if beg_chr else ".")
    right = _DELIM_MAP.get(end_chr, end_chr if end_chr else ".")
    return rf"\left{left} {inner} \right{right}"


_COLOR_WRAP_RE = re.compile(r"^\\textcolor\[HTML\]\{([0-9A-Fa-f]{6})\}\{(.*)\}$", re.DOTALL)


def _lift_color(base: str, wrapper_cmd: str) -> str:
    """`base`全体が単一の`\\textcolor{...}{...}`である場合，色指定を外側へ持ち上げる．

    `\\overline{\\textcolor{c}{x}}`のように，装飾コマンドの内側だけを着色すると，
    LaTeXの仕様上，装飾（罫線等）自体の色は着色対象に含まれず既定色（黒）の
    ままになる．装飾コマンドと文字の色を一致させるため，色指定が`base`全体を
    覆っている場合に限り，`\\textcolor{c}{\\overline{x}}`のように外側へ移す．

    引数:
        base (str): 装飾コマンドの内側に入れる予定のLaTeX文字列．
        wrapper_cmd (str): 適用する装飾（例: `r"\\overline{{{}}}"`）．`{}`に
            装飾対象の内容を埋め込む書式文字列．
    戻り値:
        str: 装飾を適用した最終的なLaTeX文字列．
    """

    match = _COLOR_WRAP_RE.match(base)
    if match is None:
        return wrapper_cmd.format(base)
    hex_color, inner = match.group(1), match.group(2)
    return rf"\textcolor[HTML]{{{hex_color}}}{{{wrapper_cmd.format(inner)}}}"


def _convert_rad(element: etree._Element, ctx: OmmlConversionContext) -> str:
    rad_pr = _find(element, "m:radPr")
    hide_deg = False
    if rad_pr is not None:
        hide_node = _find(rad_pr, "m:degHide")
        if hide_node is not None and hide_node.get("{%s}val" % _M_NS, "1") in ("1", "on", "true"):
            hide_deg = True

    base = _convert_e(element, ctx)
    deg_node = _find(element, "m:deg")
    deg = _convert_children(deg_node, ctx) if deg_node is not None else ""

    if hide_deg or not deg.strip():
        return _lift_color(base, r"\sqrt{{{}}}")
    return _lift_color(base, r"\sqrt[" + deg + r"]{{{}}}")


def _convert_nary(element: etree._Element, ctx: OmmlConversionContext) -> str:
    nary_pr = _find(element, "m:naryPr")
    op_cmd = r"\sum"
    if nary_pr is not None:
        chr_node = _find(nary_pr, "m:chr")
        if chr_node is not None:
            char = chr_node.get("{%s}val" % _M_NS, "")
            op_cmd = _NARY_CHR_MAP.get(char, char if char else op_cmd)

    sub = _convert_sub_sup_text(element, "m:sub", ctx)
    sup = _convert_sub_sup_text(element, "m:sup", ctx)
    base = _convert_e(element, ctx)

    result = op_cmd
    if sub:
        result += f"_{{{sub}}}"
    if sup:
        result += f"^{{{sup}}}"
    return f"{result} {base}"


def _convert_func(element: etree._Element, ctx: OmmlConversionContext) -> str:
    f_name_node = _find(element, "m:fName")
    name = _convert_children(f_name_node, ctx) if f_name_node is not None else ""
    base = _convert_e(element, ctx)

    plain_name = re.sub(r"[^A-Za-z]", "", name)
    if plain_name.lower() in _KNOWN_FUNC_NAMES:
        return rf"\{plain_name.lower()} {base}"
    return rf"\operatorname{{{name}}}{{{base}}}"


def _convert_limLow(element: etree._Element, ctx: OmmlConversionContext) -> str:
    base = _convert_e(element, ctx)
    lim = _convert_sub_sup_text(element, "m:lim", ctx)
    return rf"\operatorname*{{{base}}}_{{{lim}}}"


def _convert_limUpp(element: etree._Element, ctx: OmmlConversionContext) -> str:
    base = _convert_e(element, ctx)
    lim = _convert_sub_sup_text(element, "m:lim", ctx)
    return rf"\overset{{{lim}}}{{{base}}}"


def _convert_acc(element: etree._Element, ctx: OmmlConversionContext) -> str:
    acc_pr = _find(element, "m:accPr")
    cmd = r"\overline"
    if acc_pr is not None:
        chr_node = _find(acc_pr, "m:chr")
        if chr_node is not None:
            char = chr_node.get("{%s}val" % _M_NS, "")
            cmd = _ACCENT_MAP.get(char, cmd)

    base = _convert_e(element, ctx)
    return _lift_color(base, cmd + r"{{{}}}")


def _convert_bar(element: etree._Element, ctx: OmmlConversionContext) -> str:
    bar_pr = _find(element, "m:barPr")
    pos = "top"
    if bar_pr is not None:
        pos_node = _find(bar_pr, "m:pos")
        if pos_node is not None:
            pos = pos_node.get("{%s}val" % _M_NS, pos)

    base = _convert_e(element, ctx)
    cmd = r"\overline" if pos == "top" else r"\underline"
    return _lift_color(base, cmd + r"{{{}}}")


def _convert_groupChr(element: etree._Element, ctx: OmmlConversionContext) -> str:
    pr = _find(element, "m:groupChrPr")
    pos = "top"
    if pr is not None:
        pos_node = _find(pr, "m:pos")
        if pos_node is not None:
            pos = pos_node.get("{%s}val" % _M_NS, pos)

    base = _convert_e(element, ctx)
    cmd = r"\overbrace" if pos == "top" else r"\underbrace"
    return _lift_color(base, cmd + r"{{{}}}")


def _convert_eqArr(element: etree._Element, ctx: OmmlConversionContext) -> str:
    rows = [_convert_children(e, ctx) for e in element.findall(qn("m:e"))]
    body = r" \\ ".join(rows)
    return rf"\begin{{aligned}} {body} \end{{aligned}}"


def _convert_matrix(element: etree._Element, ctx: OmmlConversionContext) -> str:
    row_strs = []
    for mr in element.findall(qn("m:mr")):
        cells = [_convert_children(e, ctx) for e in mr.findall(qn("m:e"))]
        row_strs.append(" & ".join(cells))
    body = r" \\ ".join(row_strs)
    return rf"\begin{{matrix}} {body} \end{{matrix}}"


def _convert_box(element: etree._Element, ctx: OmmlConversionContext) -> str:
    return _convert_e(element, ctx)


def _convert_borderBox(element: etree._Element, ctx: OmmlConversionContext) -> str:
    base = _convert_e(element, ctx)
    return rf"\boxed{{{base}}}"


_DISPATCH = {
    qn("m:r"): _convert_run,
    qn("m:sSub"): _convert_sSub,
    qn("m:sSup"): _convert_sSup,
    qn("m:sSubSup"): _convert_sSubSup,
    qn("m:sPre"): _convert_sPre,
    qn("m:f"): _convert_f,
    qn("m:d"): _convert_d,
    qn("m:rad"): _convert_rad,
    qn("m:nary"): _convert_nary,
    qn("m:func"): _convert_func,
    qn("m:limLow"): _convert_limLow,
    qn("m:limUpp"): _convert_limUpp,
    qn("m:acc"): _convert_acc,
    qn("m:bar"): _convert_bar,
    qn("m:groupChr"): _convert_groupChr,
    qn("m:eqArr"): _convert_eqArr,
    qn("m:m"): _convert_matrix,
    qn("m:box"): _convert_box,
    qn("m:borderBox"): _convert_borderBox,
}


def _convert_node(element: etree._Element, ctx: OmmlConversionContext) -> str:
    """1つのOMML要素をLaTeX文字列へ変換する（内部再帰用のディスパッチャ）．"""

    handler = _DISPATCH.get(element.tag)
    if handler is None:
        local_tag = etree.QName(element).localname
        ctx.warn_unsupported(f"m:{local_tag}")
        return _convert_children(element, ctx)
    return handler(element, ctx)


def convert_omath_to_latex(
    omath_element: etree._Element,
    warning_log: WarningLog | None = None,
    slide_index: int | None = None,
) -> str:
    """`m:oMath`要素をLaTeXの数式本体（`$`を含まない）へ変換する．

    引数:
        omath_element (etree._Element): `m:oMath`要素．
        warning_log (WarningLog | None): 警告記録先．Noneの場合は共有ロガーを使用する．
        slide_index (int | None): 対象スライド番号（警告表示用）．
    戻り値:
        str: LaTeXの数式本体文字列．
    """

    ctx = OmmlConversionContext(warning_log=warning_log or default_log, slide_index=slide_index)
    return _convert_children(omath_element, ctx)
