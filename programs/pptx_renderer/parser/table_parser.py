"""`p:graphicFrame`（表・グラフ等）をSlide IRへ変換するパーサー．

現時点では表（`a:tbl`）のみに対応する．グラフ（chart）等は未対応要素として
警告を出した上で無視する．
"""

from __future__ import annotations

from dataclasses import dataclass

from lxml import etree

from pptx_renderer.ir import Fill, RGBColor, Stroke, TableCell, TableRow, TableShape
from pptx_renderer.parser.common import a, p, parse_xfrm, resolve_color_element
from pptx_renderer.parser.package import PptxPackage
from pptx_renderer.parser.text_parser import parse_text_body
from pptx_renderer.parser.theme import Theme
from pptx_renderer.warnings_log import WarningLog

_TABLE_STYLES_PART = "ppt/tableStyles.xml"

_TBL_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"


def _a(tag: str) -> str:
    return f"{{{_TBL_NS}}}{tag}"


_CHART_URI = "http://schemas.openxmlformats.org/drawingml/2006/chart"
_TABLE_URI = "http://schemas.openxmlformats.org/drawingml/2006/table"


@dataclass
class _TableStyle:
    """`ppt/tableStyles.xml`の1スタイル分の塗り情報（罫線・文字装飾は非対応）．"""

    whole_fill: Fill
    band1h_fill: Fill | None
    band2h_fill: Fill | None
    first_row_fill: Fill | None
    last_row_fill: Fill | None


def _parse_tcstyle_fill(tc_style_el: etree._Element | None, theme: Theme) -> Fill | None:
    """`a:tcStyle/a:fill/a:solidFill`から塗りを取得する．未指定ならNone．"""

    if tc_style_el is None:
        return None
    fill_el = tc_style_el.find(a("fill"))
    if fill_el is None:
        return None
    solid = fill_el.find(a("solidFill"))
    if solid is None:
        return None
    color_el = solid.find(a("srgbClr"))
    if color_el is None:
        color_el = solid.find(a("schemeClr"))
    if color_el is None:
        return None
    color, alpha = resolve_color_element(color_el, theme)
    return Fill(kind="solid", color=color, alpha=alpha)


def _load_table_styles(package: PptxPackage, theme: Theme) -> dict[str, _TableStyle]:
    """`ppt/tableStyles.xml`を解析し，スタイルID毎の塗り情報を返す．

    引数:
        package (PptxPackage): PPTXパッケージ．
        theme (Theme): 配色解決に使用するテーマ情報．
    戻り値:
        dict[str, _TableStyle]: スタイルIDをキーとする塗り情報の辞書．
            `ppt/tableStyles.xml`が存在しない場合は空の辞書．
    """

    if not package.exists(_TABLE_STYLES_PART):
        return {}

    root = package.read_xml(_TABLE_STYLES_PART)
    styles: dict[str, _TableStyle] = {}
    for style_el in root.findall(a("tblStyle")):
        style_id = style_el.get("styleId")
        if not style_id:
            continue
        whole_fill = _parse_tcstyle_fill(style_el.find(f"{a('wholeTbl')}/{a('tcStyle')}"), theme)
        styles[style_id] = _TableStyle(
            whole_fill=whole_fill or Fill(kind="none"),
            band1h_fill=_parse_tcstyle_fill(style_el.find(f"{a('band1H')}/{a('tcStyle')}"), theme),
            band2h_fill=_parse_tcstyle_fill(style_el.find(f"{a('band2H')}/{a('tcStyle')}"), theme),
            first_row_fill=_parse_tcstyle_fill(style_el.find(f"{a('firstRow')}/{a('tcStyle')}"), theme),
            last_row_fill=_parse_tcstyle_fill(style_el.find(f"{a('lastRow')}/{a('tcStyle')}"), theme),
        )
    return styles


def _row_default_fill(
    row_index: int, n_rows: int, table_style: _TableStyle | None, first_row_on: bool, last_row_on: bool, band_row_on: bool
) -> Fill:
    """行番号とテーブルスタイルから，セル未指定時の既定塗りを決定する．

    引数:
        row_index (int): 0始まりの行番号．
        n_rows (int): 総行数．
        table_style (_TableStyle | None): 適用するテーブルスタイル．
        first_row_on (bool): `tblPr/@firstRow`が有効か．
        last_row_on (bool): `tblPr/@lastRow`が有効か．
        band_row_on (bool): `tblPr/@bandRow`が有効か．
    戻り値:
        Fill: 既定の塗り（テーブルスタイルが無い場合は`Fill(kind="none")`）．
    """

    if table_style is None:
        return Fill(kind="none")

    if first_row_on and row_index == 0 and table_style.first_row_fill is not None:
        return table_style.first_row_fill
    if last_row_on and row_index == n_rows - 1 and table_style.last_row_fill is not None:
        return table_style.last_row_fill

    if band_row_on:
        band_start = 1 if first_row_on else 0
        band_offset = row_index - band_start
        if band_offset >= 0:
            banded = table_style.band1h_fill if band_offset % 2 == 0 else table_style.band2h_fill
            return banded if banded is not None else table_style.whole_fill

    return table_style.whole_fill


def parse_graphic_frame(
    gf_el: etree._Element, theme: Theme, package: PptxPackage, warning_log: WarningLog, slide_index: int | None
) -> TableShape | None:
    """`p:graphicFrame`を`TableShape`へ変換する．

    引数:
        gf_el (etree._Element): `p:graphicFrame`要素．
        theme (Theme): テーマ情報．
        package (PptxPackage): テーブルスタイル（`ppt/tableStyles.xml`）解決用のパッケージ．
        warning_log (WarningLog): 警告記録先．
        slide_index (int | None): 対象スライド番号．
    戻り値:
        TableShape | None: 表として解釈できた場合のみ変換結果を返す．
    """

    c_nv_pr = gf_el.find(f"{p('nvGraphicFramePr')}/{p('cNvPr')}")
    shape_id = c_nv_pr.get("id", "0") if c_nv_pr is not None else "0"
    name = c_nv_pr.get("name", "") if c_nv_pr is not None else ""

    xfrm = gf_el.find(p("xfrm"))
    rect, rotation, _, _ = parse_xfrm(xfrm)

    graphic_data = gf_el.find(f"{a('graphic')}/{a('graphicData')}")
    if graphic_data is None:
        warning_log.unsupported_element("p:graphicFrame", slide_index)
        return None

    uri = graphic_data.get("uri", "")
    if uri == _CHART_URI:
        warning_log.add("unsupported_element", "グラフ（chart）は未対応のため描画をスキップしました．", slide_index)
        return None
    if uri != _TABLE_URI:
        warning_log.unsupported_element(f"p:graphicFrame (uri={uri})", slide_index)
        return None

    tbl_el = graphic_data.find(a("tbl"))
    if tbl_el is None:
        return None

    col_widths = [float(gc.get("w")) for gc in tbl_el.findall(f"{a('tblGrid')}/{a('gridCol')}")]

    tbl_pr = tbl_el.find(a("tblPr"))
    first_row_on = tbl_pr is not None and tbl_pr.get("firstRow") == "1"
    last_row_on = tbl_pr is not None and tbl_pr.get("lastRow") == "1"
    band_row_on = tbl_pr is not None and tbl_pr.get("bandRow") == "1"
    style_id_el = tbl_pr.find(a("tableStyleId")) if tbl_pr is not None else None
    style_id = style_id_el.text.strip() if style_id_el is not None and style_id_el.text else None
    table_style = _load_table_styles(package, theme).get(style_id) if style_id else None

    tr_els = tbl_el.findall(a("tr"))
    rows: list[TableRow] = []
    for row_index, tr_el in enumerate(tr_els):
        height = float(tr_el.get("h", "0"))
        default_fill = _row_default_fill(row_index, len(tr_els), table_style, first_row_on, last_row_on, band_row_on)
        cells = [_parse_tc(tc_el, theme, warning_log, slide_index, default_fill) for tc_el in tr_el.findall(a("tc"))]
        rows.append(TableRow(height_emu=height, cells=cells))

    return TableShape(id=shape_id, name=name, rect=rect, col_widths_emu=col_widths, rows=rows, rotation=rotation)


def _parse_border(tc_pr: etree._Element, tag: str, theme: Theme) -> Stroke:
    ln_el = tc_pr.find(a(tag))
    if ln_el is None:
        return Stroke(kind="none")
    if ln_el.find(a("noFill")) is not None:
        return Stroke(kind="none")
    solid = ln_el.find(a("solidFill"))
    if solid is None:
        return Stroke(kind="none")
    color_el = solid.find(a("srgbClr"))
    if color_el is None:
        color_el = solid.find(a("schemeClr"))
    if color_el is None:
        return Stroke(kind="none")
    color, alpha = resolve_color_element(color_el, theme)
    width_pt = float(ln_el.get("w", "12700")) / 12700.0
    return Stroke(kind="solid", color=color, width_pt=width_pt, alpha=alpha)


def _parse_tc(
    tc_el: etree._Element, theme: Theme, warning_log: WarningLog, slide_index: int | None, default_fill: Fill
) -> TableCell:
    """`a:tc`（表のセル）を`TableCell`へ変換する．

    引数:
        tc_el (etree._Element): `a:tc`要素．
        theme (Theme): テーマ情報．
        warning_log (WarningLog): 警告記録先．
        slide_index (int | None): 対象スライド番号．
        default_fill (Fill): セル自身に塗りの指定が無い場合に用いる既定の塗り
            （テーブルスタイルの縞模様・見出し行等から算出される）．
    戻り値:
        TableCell: 変換結果．
    """

    is_covered = tc_el.get("hMerge") == "1" or tc_el.get("vMerge") == "1"
    col_span = int(tc_el.get("gridSpan", "1"))
    row_span = int(tc_el.get("rowSpan", "1"))

    tc_pr = tc_el.find(a("tcPr"))
    fill = Fill(kind="none")
    has_explicit_no_fill = False
    border_left = border_right = border_top = border_bottom = Stroke(kind="none")
    if tc_pr is not None:
        has_explicit_no_fill = tc_pr.find(a("noFill")) is not None
        solid = tc_pr.find(a("solidFill"))
        if solid is not None:
            color_el = solid.find(a("srgbClr"))
            if color_el is None:
                color_el = solid.find(a("schemeClr"))
            if color_el is not None:
                color, alpha = resolve_color_element(color_el, theme)
                fill = Fill(kind="solid", color=color, alpha=alpha)
        border_left = _parse_border(tc_pr, "lnL", theme)
        border_right = _parse_border(tc_pr, "lnR", theme)
        border_top = _parse_border(tc_pr, "lnT", theme)
        border_bottom = _parse_border(tc_pr, "lnB", theme)

    if fill.kind == "none" and not has_explicit_no_fill:
        fill = default_fill

    tx_body_el = tc_el.find(a("txBody"))
    text_body = parse_text_body(tx_body_el, theme, warning_log, slide_index, default_size_pt=18.0)

    return TableCell(
        text_body=text_body,
        fill=fill,
        col_span=col_span,
        row_span=row_span,
        is_covered=is_covered,
        border_left=border_left,
        border_right=border_right,
        border_top=border_top,
        border_bottom=border_bottom,
    )
