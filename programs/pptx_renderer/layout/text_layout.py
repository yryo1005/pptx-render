"""テキストの折り返し（wrapping）・行送り計算を行うモジュール．

`text box`，`paragraph`，`run`，`font metrics`，`line spacing`，`wrapping`を
考慮して，PowerPointの改行位置・行間にできるだけ近いレイアウトを計算する．
日本語は原則として文字単位で折り返し，禁則処理（行頭禁則・行末禁則）の
簡易版を適用する．英数字は単語単位で折り返す．
"""

from __future__ import annotations

from dataclasses import dataclass, field

from reportlab.pdfbase import pdfmetrics

from pptx_renderer.fonts.registry import FontRegistry, ResolvedFont
from pptx_renderer.ir import HAlign, LineBreak, MathRun, Paragraph, RGBColor, TextBody, TextRun
from pptx_renderer.math.latex_render import LatexMathRenderer, MathRenderResult
from pptx_renderer.warnings_log import WarningLog

# 行頭に置いてはならない文字（前の行の末尾へ送る）．
_FORBIDDEN_LEADING = set("、。，．・：；？！ー）］｝」』〉》’”ゝゞぁぃぅぇぉっゃゅょゎァィゥェォッャュョヮ%),.:;?!'\">」』】")
# 行末に置いてはならない文字（次の行の先頭へ送る）．
_FORBIDDEN_TRAILING = set("（［｛「『〈《‘“(['\"「【")

_CJK_RANGES = (
    (0x3040, 0x30FF),
    (0x3400, 0x4DBF),
    (0x4E00, 0x9FFF),
    (0xF900, 0xFAFF),
    (0xFF00, 0xFFEF),
)


def _is_cjk(ch: str) -> bool:
    code = ord(ch)
    return any(lo <= code <= hi for lo, hi in _CJK_RANGES)


@dataclass
class TextAtom:
    """折り返し単位となる最小のテキスト片．"""

    text: str
    is_space: bool = False


def tokenize_for_wrap(text: str) -> list[TextAtom]:
    """文字列を折り返し単位（アトム）へ分割する．

    引数:
        text (str): 対象文字列．
    戻り値:
        list[TextAtom]: アトムのリスト．CJK文字は1文字ずつ，非CJK文字は
            空白区切りの単語単位でアトム化する．
    """

    atoms: list[TextAtom] = []
    buffer = ""
    for ch in text:
        if ch.isspace():
            if buffer:
                atoms.append(TextAtom(buffer))
                buffer = ""
            atoms.append(TextAtom(ch, is_space=True))
        elif _is_cjk(ch):
            if buffer:
                atoms.append(TextAtom(buffer))
                buffer = ""
            atoms.append(TextAtom(ch))
        else:
            buffer += ch
    if buffer:
        atoms.append(TextAtom(buffer))
    return atoms


@dataclass
class Segment:
    """1行内の描画単位（テキストラン片または数式）．"""

    kind: str  # "text" | "math"
    text: str = ""
    font_name: str = ""
    reportlab_font: str = ""
    size_pt: float = 12.0
    bold: bool = False
    italic: bool = False
    underline: bool = False
    color: RGBColor = field(default_factory=lambda: RGBColor(0, 0, 0))
    alpha: float = 1.0
    width_pt: float = 0.0
    ascent_pt: float = 0.0
    descent_pt: float = 0.0
    math_result: MathRenderResult | None = None
    needs_faux_bold: bool = False
    needs_faux_italic: bool = False
    char_spacing_pt: float = 0.0


@dataclass
class LineBox:
    """レイアウト済みの1行分の情報．"""

    segments: list[Segment] = field(default_factory=list)
    align: HAlign = HAlign.LEFT
    line_height_pt: float = 0.0
    ascent_pt: float = 0.0
    descent_pt: float = 0.0
    space_before_pt: float = 0.0
    total_width_pt: float = 0.0
    indent_left_pt: float = 0.0
    bullet_char: str | None = None
    bullet_font: str | None = None
    bullet_size_pt: float = 12.0
    bullet_offset_pt: float = 0.0


def _font_ascent_descent(resolved: ResolvedFont, font_registry: FontRegistry, size_pt: float) -> tuple[float, float]:
    try:
        tt = font_registry.get_ttfont(resolved)
        units_per_em = tt["head"].unitsPerEm
        os2 = tt["OS/2"]
        ascent = os2.sTypoAscender / units_per_em * size_pt
        descent = -os2.sTypoDescender / units_per_em * size_pt
        return ascent, descent
    except Exception:  # noqa: BLE001
        return size_pt * 0.8, size_pt * 0.2


def _build_segment_for_run_part(
    text: str,
    run: TextRun,
    font_registry: FontRegistry,
    slide_index: int | None,
) -> Segment:
    is_cjk_atom = any(_is_cjk(ch) for ch in text)
    font_name = (run.ea_font_name if is_cjk_atom and run.ea_font_name else None) or run.font_name or "Calibri"
    resolved = font_registry.resolve(font_name, run.bold, run.italic, slide_index)
    reportlab_name = font_registry.register_reportlab_font(resolved)
    width_pt = pdfmetrics.stringWidth(text, reportlab_name, run.size_pt) + run.char_spacing_pt * len(text)
    ascent, descent = _font_ascent_descent(resolved, font_registry, run.size_pt)
    return Segment(
        kind="text",
        text=text,
        font_name=font_name,
        reportlab_font=reportlab_name,
        size_pt=run.size_pt,
        bold=run.bold,
        italic=run.italic,
        underline=run.underline,
        color=run.color,
        alpha=run.alpha,
        width_pt=width_pt,
        ascent_pt=ascent,
        descent_pt=descent,
        char_spacing_pt=run.char_spacing_pt,
        needs_faux_bold=resolved.needs_faux_bold,
        needs_faux_italic=resolved.needs_faux_italic,
    )


def _build_segment_for_math(
    run: MathRun, math_renderer: LatexMathRenderer, warning_log: WarningLog, slide_index: int | None
) -> Segment | None:
    result = math_renderer.render(run.latex_body, run.size_pt, display=run.display, slide_index=slide_index)
    if result is None:
        return None
    return Segment(
        kind="math",
        size_pt=run.size_pt,
        width_pt=result.width_pt,
        ascent_pt=result.height_pt,
        descent_pt=result.depth_pt,
        math_result=result,
    )


def _line_height_for(paragraph: Paragraph, max_size_pt: float) -> float:
    if paragraph.line_spacing_pt is not None:
        return paragraph.line_spacing_pt
    pct = paragraph.line_spacing_pct if paragraph.line_spacing_pct is not None else 100.0
    return max_size_pt * 1.2 * (pct / 100.0)


def layout_paragraph(
    paragraph: Paragraph,
    available_width_pt: float,
    font_registry: FontRegistry,
    math_renderer: LatexMathRenderer,
    warning_log: WarningLog,
    slide_index: int | None,
) -> list[LineBox]:
    """1段落分の折り返しを計算し，行（LineBox）のリストを返す．

    引数:
        paragraph (Paragraph): レイアウト対象の段落．
        available_width_pt (float): 折り返しの基準となる利用可能幅（pt）．
        font_registry (FontRegistry): フォント解決用レジストリ．
        math_renderer (LatexMathRenderer): 数式レンダラー．
        warning_log (WarningLog): 警告記録先．
        slide_index (int | None): 対象スライド番号．
    戻り値:
        list[LineBox]: レイアウト済みの行のリスト（段落が空の場合も最低1行を返す）．
    """

    indent_left_pt = paragraph.indent_left_emu / 12700.0
    wrap_width_pt = max(1.0, available_width_pt - indent_left_pt)
    empty_fallback_size = paragraph.empty_line_size_pt or 12.0

    lines: list[LineBox] = []
    current: list[Segment] = []
    current_width = 0.0

    max_size_for_line = 1.0
    # 自動折り返し（幅超過）による改行の直後にのみ，行頭の空白を抑制する．
    # 段落の先頭や明示的な改行の直後は，PPTX側で意図的に挿入された
    # 字下げ用の空白（擬似コードのインデント等）である場合があるため抑制しない．
    just_auto_wrapped = False

    def flush_line(force_empty: bool = False, auto_wrap: bool = False) -> None:
        nonlocal current, current_width, max_size_for_line, just_auto_wrapped
        just_auto_wrapped = auto_wrap
        if not current and not force_empty:
            return
        fallback_size = max_size_for_line if current else empty_fallback_size
        ascent = max((s.ascent_pt for s in current), default=fallback_size * 0.8)
        descent = max((s.descent_pt for s in current), default=fallback_size * 0.2)
        # 数式（特にsum/limit付きの表示数式）は，名目上のフォントサイズから
        # 期待される行送りより実際の高さがはるかに大きくなる場合があるため，
        # 行送りは名目値と実際のascent+descentの大きい方を採用する．
        line_height = max(_line_height_for(paragraph, fallback_size), ascent + descent)
        lines.append(
            LineBox(
                segments=current,
                align=paragraph.align,
                line_height_pt=line_height,
                ascent_pt=ascent,
                descent_pt=descent,
                total_width_pt=current_width,
                indent_left_pt=indent_left_pt,
            )
        )
        current = []
        current_width = 0.0
        max_size_for_line = 1.0

    for run in paragraph.runs:
        if isinstance(run, LineBreak):
            flush_line(force_empty=True)
            continue

        if isinstance(run, MathRun):
            seg = _build_segment_for_math(run, math_renderer, warning_log, slide_index)
            if seg is None:
                continue
            if current_width + seg.width_pt > wrap_width_pt and current:
                flush_line(auto_wrap=True)
            current.append(seg)
            current_width += seg.width_pt
            max_size_for_line = max(max_size_for_line, run.size_pt)
            just_auto_wrapped = False
            continue

        if isinstance(run, TextRun):
            if run.text == "":
                continue
            atoms = tokenize_for_wrap(run.text)
            i = 0
            while i < len(atoms):
                atom = atoms[i]
                if atom.is_space:
                    if current or not just_auto_wrapped:
                        seg = _build_segment_for_run_part(atom.text, run, font_registry, slide_index)
                        current.append(seg)
                        current_width += seg.width_pt
                        max_size_for_line = max(max_size_for_line, run.size_pt)
                        just_auto_wrapped = False
                    i += 1
                    continue

                seg = _build_segment_for_run_part(atom.text, run, font_registry, slide_index)

                if current_width + seg.width_pt > wrap_width_pt and current:
                    if atom.text and atom.text[0] in _FORBIDDEN_LEADING:
                        current.append(seg)
                        current_width += seg.width_pt
                        max_size_for_line = max(max_size_for_line, run.size_pt)
                        i += 1
                        continue
                    flush_line(auto_wrap=True)

                current.append(seg)
                current_width += seg.width_pt
                max_size_for_line = max(max_size_for_line, run.size_pt)
                just_auto_wrapped = False
                i += 1

    flush_line(force_empty=len(lines) == 0)

    if lines:
        lines[0].space_before_pt = paragraph.space_before_pt
        if paragraph.bullet_char and lines[0].segments:
            lines[0].bullet_char = paragraph.bullet_char
            # 自動採番（1. や a) 等）は本文と同じフォントで描画するのが自然なため，
            # 記号用フォント（buFont，既定Arial）は文字（buChar）の場合のみ使う．
            lines[0].bullet_font = lines[0].segments[0].font_name if paragraph.bullet_auto_num_fmt else paragraph.bullet_font
            lines[0].bullet_size_pt = lines[0].segments[0].size_pt
            lines[0].bullet_offset_pt = paragraph.bullet_offset_emu / 12700.0

    return lines


def layout_text_body(
    text_body: TextBody,
    available_width_pt: float,
    font_registry: FontRegistry,
    math_renderer: LatexMathRenderer,
    warning_log: WarningLog,
    slide_index: int | None,
) -> list[LineBox]:
    """テキスト本体全体（複数段落）のレイアウトを計算する．

    引数:
        text_body (TextBody): レイアウト対象のテキスト本体．
        available_width_pt (float): 折り返しの基準となる利用可能幅（pt）．
        font_registry (FontRegistry): フォント解決用レジストリ．
        math_renderer (LatexMathRenderer): 数式レンダラー．
        warning_log (WarningLog): 警告記録先．
        slide_index (int | None): 対象スライド番号．
    戻り値:
        list[LineBox]: 全段落を通した行のリスト．
    """

    all_lines: list[LineBox] = []
    for paragraph in text_body.paragraphs:
        width = available_width_pt if text_body.wrap else float("inf")
        all_lines.extend(
            layout_paragraph(paragraph, width, font_registry, math_renderer, warning_log, slide_index)
        )
        if all_lines:
            all_lines[-1].line_height_pt += paragraph.space_after_pt
    return all_lines
