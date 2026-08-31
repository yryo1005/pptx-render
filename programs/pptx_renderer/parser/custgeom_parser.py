"""`a:custGeom`（カスタムベクター図形）の`a:pathLst`をSlide IRの
`CustomPath`（描画コマンド列）へ変換するパーサー．

対応するパスコマンド: `a:moveTo`, `a:lnTo`, `a:cubicBezTo`, `a:quadBezTo`,
`a:arcTo`（円弧を折れ線で近似），`a:close`．

`a:gdLst`（名前付きガイド）はパス座標の直接参照には用いられないのが
一般的であるため，本パーサーではパス上の座標（`a:pt/@x`, `@y`）を
リテラルな数値としてのみ解釈する．
"""

from __future__ import annotations

import math

from lxml import etree

from pptx_renderer.ir import CustomPath
from pptx_renderer.parser.common import a


def parse_cust_geom(cust_geom_el: etree._Element) -> list[CustomPath]:
    """`a:custGeom`要素から`CustomPath`のリストを生成する．

    引数:
        cust_geom_el (etree._Element): `a:custGeom`要素．
    戻り値:
        list[CustomPath]: パスごとの描画コマンド列．`a:pathLst`が無い場合は空リスト．
    """

    path_lst = cust_geom_el.find(a("pathLst"))
    if path_lst is None:
        return []

    return [_parse_path(path_el) for path_el in path_lst.findall(a("path"))]


def _pt(el: etree._Element) -> tuple[float, float]:
    pt_el = el.find(a("pt"))
    return float(pt_el.get("x")), float(pt_el.get("y"))


def _parse_path(path_el: etree._Element) -> CustomPath:
    width = float(path_el.get("w", "1"))
    height = float(path_el.get("h", "1"))
    fill = path_el.get("fill", "norm") != "none"
    stroke = path_el.get("stroke", "1") != "0"

    commands: list[tuple] = []
    current = (0.0, 0.0)

    for el in path_el:
        local = etree.QName(el).localname
        if local == "moveTo":
            current = _pt(el)
            commands.append(("moveTo", *current))
        elif local == "lnTo":
            current = _pt(el)
            commands.append(("lineTo", *current))
        elif local == "cubicBezTo":
            pts = el.findall(a("pt"))
            c1 = (float(pts[0].get("x")), float(pts[0].get("y")))
            c2 = (float(pts[1].get("x")), float(pts[1].get("y")))
            end = (float(pts[2].get("x")), float(pts[2].get("y")))
            commands.append(("curveTo", *c1, *c2, *end))
            current = end
        elif local == "quadBezTo":
            pts = el.findall(a("pt"))
            q = (float(pts[0].get("x")), float(pts[0].get("y")))
            end = (float(pts[1].get("x")), float(pts[1].get("y")))
            c1 = (current[0] + 2.0 / 3.0 * (q[0] - current[0]), current[1] + 2.0 / 3.0 * (q[1] - current[1]))
            c2 = (end[0] + 2.0 / 3.0 * (q[0] - end[0]), end[1] + 2.0 / 3.0 * (q[1] - end[1]))
            commands.append(("curveTo", *c1, *c2, *end))
            current = end
        elif local == "arcTo":
            current = _append_arc(commands, current, el)
        elif local == "close":
            commands.append(("close",))

    return CustomPath(width_emu=width, height_emu=height, commands=commands, fill=fill, stroke=stroke)


def _append_arc(commands: list[tuple], current: tuple[float, float], arc_el: etree._Element) -> tuple[float, float]:
    """`a:arcTo`を，現在点を起点とする楕円弧として折れ線近似し，`commands`へ追記する．"""

    w_r = float(arc_el.get("wR", "0"))
    h_r = float(arc_el.get("hR", "0"))
    st_ang = float(arc_el.get("stAng", "0")) / 60000.0
    sw_ang = float(arc_el.get("swAng", "0")) / 60000.0

    if w_r == 0 or h_r == 0:
        return current

    # OOXMLのarcToは，現在点が「開始角度における楕円弧上の点」となるように
    # 楕円の中心を逆算する（現在点はそのまま弧の始点として保持される）．
    start_rad = math.radians(st_ang)
    center_x = current[0] - w_r * math.cos(start_rad)
    center_y = current[1] - h_r * math.sin(start_rad)

    steps = max(4, int(abs(sw_ang) / 10.0) + 1)
    last = current
    for i in range(1, steps + 1):
        ang = math.radians(st_ang + sw_ang * i / steps)
        pt = (center_x + w_r * math.cos(ang), center_y + h_r * math.sin(ang))
        commands.append(("lineTo", *pt))
        last = pt

    return last
