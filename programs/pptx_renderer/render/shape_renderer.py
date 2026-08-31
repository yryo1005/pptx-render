"""既定図形（rectangle, ellipse, line, arrow, polygon等）のベクター描画を行うモジュール．

PPTXの`prstGeom`（既定図形プリセット）を可能な限りPDFのベクターパスとして
再現する．未対応のプリセットは矩形として近似し，警告を記録する．
図形の回転・反転は，このモジュールが提供する`apply_shape_transform`を通じて
一元的に適用する．
"""

from __future__ import annotations

import math

from reportlab.pdfgen import canvas as canvas_module

from pptx_renderer.ir import AutoShape, ShapeStyle
from pptx_renderer.units import RectPt
from pptx_renderer.warnings_log import WarningLog


def apply_shape_transform(
    c: canvas_module.Canvas, rect_pt: RectPt, rotation_deg: float, flip_h: bool, flip_v: bool
) -> None:
    """図形の回転・反転・平行移動をキャンバスへ適用する．

    呼び出し後，図形はローカル座標系（左下原点(0,0)，右上(width,height)）で
    描画できる．呼び出し前後で`canvas.saveState()`/`restoreState()`を
    対にして使用すること．

    引数:
        c (canvas_module.Canvas): 描画対象のキャンバス．
        rect_pt (RectPt): PDF座標系での図形の矩形（左下原点）．
        rotation_deg (float): 時計回りの回転角度（度）．
        flip_h (bool): 水平反転するか．
        flip_v (bool): 垂直反転するか．
    戻り値:
        なし．
    """

    cx = rect_pt.x + rect_pt.width / 2.0
    cy = rect_pt.y + rect_pt.height / 2.0
    c.translate(cx, cy)
    if rotation_deg:
        c.rotate(-rotation_deg)
    c.scale(-1 if flip_h else 1, -1 if flip_v else 1)
    c.translate(-rect_pt.width / 2.0, -rect_pt.height / 2.0)


def _set_fill_stroke(c: canvas_module.Canvas, style: ShapeStyle) -> tuple[bool, bool]:
    do_fill = style.fill.kind == "solid" and style.fill.color is not None
    do_stroke = style.stroke.kind == "solid" and style.stroke.color is not None

    if do_fill:
        c.setFillColorRGB(*style.fill.color.to_unit_tuple())
        c.setFillAlpha(style.fill.alpha)
    if do_stroke:
        c.setStrokeColorRGB(*style.stroke.color.to_unit_tuple())
        c.setStrokeAlpha(style.stroke.alpha)
        c.setLineWidth(style.stroke.width_pt)
        if style.stroke.dash == "dash":
            c.setDash([style.stroke.width_pt * 3, style.stroke.width_pt * 2])
        elif style.stroke.dash == "dashDot":
            c.setDash([style.stroke.width_pt * 3, style.stroke.width_pt * 2, style.stroke.width_pt, style.stroke.width_pt * 2])
        elif style.stroke.dash == "dot":
            c.setDash([style.stroke.width_pt, style.stroke.width_pt * 2])
        else:
            c.setDash([])

    return do_fill, do_stroke


def _draw_polygon(c: canvas_module.Canvas, points: list[tuple[float, float]], do_fill: bool, do_stroke: bool) -> None:
    path = c.beginPath()
    path.moveTo(*points[0])
    for pt in points[1:]:
        path.lineTo(*pt)
    path.close()
    c.drawPath(path, fill=do_fill, stroke=do_stroke)


def _regular_polygon_points(w: float, h: float, n: int, start_angle_deg: float = 90.0) -> list[tuple[float, float]]:
    cx, cy = w / 2.0, h / 2.0
    rx, ry = w / 2.0, h / 2.0
    points = []
    for i in range(n):
        angle = math.radians(start_angle_deg + 360.0 * i / n)
        points.append((cx + rx * math.cos(angle), cy + ry * math.sin(angle)))
    return points


def _star_points(w: float, h: float, n: int = 5, inner_ratio: float = 0.382) -> list[tuple[float, float]]:
    cx, cy = w / 2.0, h / 2.0
    rx, ry = w / 2.0, h / 2.0
    points = []
    for i in range(n * 2):
        angle = math.radians(90.0 + 360.0 * i / (n * 2))
        ratio = 1.0 if i % 2 == 0 else inner_ratio
        points.append((cx + rx * ratio * math.cos(angle), cy + ry * ratio * math.sin(angle)))
    return points


def _block_arrow_points(w: float, h: float, direction: str) -> list[tuple[float, float]]:
    """右向き矢印を基準に，指定方向のブロック矢印の頂点列を生成する．"""

    if direction in ("left", "right"):
        head_len = min(w * 0.4, h)
        stem_top, stem_bottom = h * 0.75, h * 0.25
        pts = [
            (0, stem_bottom),
            (w - head_len, stem_bottom),
            (w - head_len, 0),
            (w, h * 0.5),
            (w - head_len, h),
            (w - head_len, stem_top),
            (0, stem_top),
        ]
        if direction == "left":
            pts = [(w - x, y) for x, y in pts]
        return pts

    head_len = min(h * 0.4, w)
    stem_left, stem_right = w * 0.25, w * 0.75
    pts = [
        (stem_left, 0),
        (stem_left, h - head_len),
        (0, h - head_len),
        (w * 0.5, h),
        (w, h - head_len),
        (stem_right, h - head_len),
        (stem_right, 0),
    ]
    if direction == "down":
        pts = [(x, h - y) for x, y in pts]
    return pts


def _double_block_arrow_points(w: float, h: float, axis: str) -> list[tuple[float, float]]:
    """`leftRightArrow`（左右）・`upDownArrow`（上下）の両端矢印の頂点列を生成する．"""

    if axis == "horizontal":
        head_len = min(w * 0.25, h)
        stem_top, stem_bottom = h * 0.75, h * 0.25
        return [
            (head_len, stem_bottom),
            (head_len, 0),
            (0, h * 0.5),
            (head_len, h),
            (head_len, stem_top),
            (w - head_len, stem_top),
            (w - head_len, h),
            (w, h * 0.5),
            (w - head_len, 0),
            (w - head_len, stem_bottom),
        ]

    head_len = min(h * 0.25, w)
    stem_left, stem_right = w * 0.25, w * 0.75
    return [
        (stem_left, head_len),
        (0, head_len),
        (w * 0.5, 0),
        (w, head_len),
        (stem_right, head_len),
        (stem_right, h - head_len),
        (w, h - head_len),
        (w * 0.5, h),
        (0, h - head_len),
        (stem_left, h - head_len),
    ]


def _draw_brace(
    c: canvas_module.Canvas, w: float, h: float, do_fill: bool, do_stroke: bool, mirror: bool
) -> None:
    """波括弧（`leftBrace`/`rightBrace`）を4本のベジェ曲線で近似描画する．

    `rightBrace`（"}"）を基準に，先端（凹みの頂点）が図形の左端(x=0)，
    背（丸みを帯びた両端）が右端(x=w)に来る形状を描く．OOXMLの調整値
    （`adj1`＝丸みの深さ，`adj2`＝先端の垂直位置）は考慮せず，固定比率で
    近似する．`leftBrace`は`mirror=True`で左右反転して描画する．
    """

    tip_x, back_x = 0.0, w
    if mirror:
        tip_x, back_x = w, 0.0

    mid_y = h * 0.5
    notch_x = back_x + (tip_x - back_x) * 0.45
    notch_gap = h * 0.06

    waypoints = [
        (back_x, h),
        (notch_x, mid_y + notch_gap),
        (tip_x, mid_y),
        (notch_x, mid_y - notch_gap),
        (back_x, 0.0),
    ]
    _draw_smooth_path(c, waypoints, do_fill, do_stroke)


def _draw_custom_paths(
    c: canvas_module.Canvas,
    custom_paths: list,
    w: float,
    h: float,
    do_fill: bool,
    do_stroke: bool,
) -> None:
    """`CustomPath`のリストを，シェイプのボックス(0,0)-(w,h)に合わせて描画する．"""

    for custom_path in custom_paths:
        scale_x = (w / custom_path.width_emu) if custom_path.width_emu else 1.0
        scale_y = (h / custom_path.height_emu) if custom_path.height_emu else 1.0

        def to_local(x: float, y: float) -> tuple[float, float]:
            return x * scale_x, h - y * scale_y

        path = c.beginPath()
        has_current_point = False
        for cmd in custom_path.commands:
            if cmd[0] == "moveTo":
                path.moveTo(*to_local(cmd[1], cmd[2]))
                has_current_point = True
            elif cmd[0] == "lineTo" and has_current_point:
                path.lineTo(*to_local(cmd[1], cmd[2]))
            elif cmd[0] == "curveTo" and has_current_point:
                x1, y1 = to_local(cmd[1], cmd[2])
                x2, y2 = to_local(cmd[3], cmd[4])
                x3, y3 = to_local(cmd[5], cmd[6])
                path.curveTo(x1, y1, x2, y2, x3, y3)
            elif cmd[0] == "close":
                path.close()

        path_fill = do_fill and custom_path.fill
        path_stroke = do_stroke and custom_path.stroke
        if path_fill or path_stroke:
            c.drawPath(path, fill=path_fill, stroke=path_stroke)


def _callout_points(w: float, h: float, adjustments: dict) -> list[tuple[float, float]]:
    """吹き出し系プリセット（wedgeRectCallout等）の,本体＋引き出し線を近似した多角形を生成する．"""

    adj1 = adjustments.get("adj1", -20833) / 100000.0
    adj2 = adjustments.get("adj2", 62500) / 100000.0

    tip_x = adj1 * w
    tip_y = h - adj2 * h  # OOXMLはY軸下向きのため反転

    bl, br, tr, tl = (0.0, 0.0), (w, 0.0), (w, h), (0.0, h)
    corners = [bl, br, tr, tl]

    if tip_y < 0:
        edge, idx = "bottom", 0
    elif tip_y > h:
        edge, idx = "top", 2
    elif tip_x < 0:
        edge, idx = "left", 3
    elif tip_x > w:
        edge, idx = "right", 1
    else:
        edge, idx = "bottom", 0

    # OOXMLの正確な引き出し線公式は簡略化しているため，adjの値が極端な場合に
    # 引き出し線が本体から大きく離れた図形・文字と重なるのを避けるよう，
    # 箱からの突出距離を短辺の35%を上限に制限する．
    max_extent = min(w, h) * 0.35
    if edge == "bottom":
        tip_y = max(tip_y, -max_extent)
    elif edge == "top":
        tip_y = min(tip_y, h + max_extent)
    elif edge == "left":
        tip_x = max(tip_x, -max_extent)
    elif edge == "right":
        tip_x = min(tip_x, w + max_extent)

    if edge in ("bottom", "top"):
        center = min(max(tip_x, 0.0), w)
        spread = w * 0.12
        p1 = (min(max(center - spread, 0.0), w), 0.0 if edge == "bottom" else h)
        p2 = (min(max(center + spread, 0.0), w), 0.0 if edge == "bottom" else h)
    else:
        center = min(max(tip_y, 0.0), h)
        spread = h * 0.12
        p1 = (0.0 if edge == "left" else w, min(max(center - spread, 0.0), h))
        p2 = (0.0 if edge == "left" else w, min(max(center + spread, 0.0), h))

    corner_a = corners[idx]

    def dist2(p, q):
        return (p[0] - q[0]) ** 2 + (p[1] - q[1]) ** 2

    near_a, near_b = (p1, p2) if dist2(p1, corner_a) <= dist2(p2, corner_a) else (p2, p1)
    tail_sequence = [near_a, (tip_x, tip_y), near_b]

    polygon: list[tuple[float, float]] = []
    for k in range(4):
        polygon.append(corners[k])
        if k == idx:
            polygon.extend(tail_sequence)
    return polygon


def _connector_points(
    w: float,
    h: float,
    preset: str,
    start_idx: int | None = None,
    end_idx: int | None = None,
) -> list[tuple[float, float]]:
    """コネクタの始点・（屈曲/制御）点・終点を，ローカルボックス(0,0)-(w,h)基準で求める．

    始点は左上(0,h)，終点は右下(w,0)とする（反転・回転は呼び出し側の
    `apply_shape_transform`が担う）．

    `a:stCxn`/`a:endCxn`の`idx`（接続先図形上の接続点番号）が分かる場合は，
    既定の接続点定義（0=上，1=左，2=下，3=右）に基づき，コネクタが図形から
    垂直に出る辺か水平に出る辺かを判定し，その方向にまず経路を伸ばしてから
    折れ曲がる（PowerPointの既定のカギ線経路に近い挙動）．情報が無い場合は
    幅方向中央で1回折れるZ字型の経路を既定として使用する．

    ただし，`w`・`h`のどちらかが極端に小さい（回転済みの細長いコネクタ等，
    ローカルなボックスの短辺がほぼ0の）場合は，Z字経路の折れ幅がほぼ0になり
    かえって鋭角な折れ曲がりに見えてしまうため，この場合は始点・終点を
    直接結ぶ単純な経路にフォールバックする．
    """

    start = (0.0, h)
    end = (w, 0.0)
    if not (preset.startswith("bent") or preset.startswith("curved")):
        return [start, end]
    if min(w, h) < 4.0:
        return [start, end]

    start_vertical = start_idx in (0, 2) if start_idx is not None else None
    end_vertical = end_idx in (0, 2) if end_idx is not None else None

    if start_vertical is not None and end_vertical is not None and start_vertical != end_vertical:
        # 片方が縦方向，もう片方が横方向に出る場合は，1回だけ折れるL字経路．
        if start_vertical:
            bend = (start[0], end[1])
        else:
            bend = (end[0], start[1])
        return [start, bend, end]

    if start_vertical is True and end_vertical is True:
        # 両方とも縦方向（上下）に出る場合は，高さ方向中央を通るZ字経路．
        mid_y = (start[1] + end[1]) / 2.0
        return [start, (start[0], mid_y), (end[0], mid_y), end]

    # 情報が無い，または両方とも横方向（左右）に出る場合は，
    # 幅方向中央を通るZ字経路を既定とする．
    mid_x = (start[0] + end[0]) / 2.0
    return [start, (mid_x, start[1]), (mid_x, end[1]), end]


def _midpoint(p1: tuple[float, float], p2: tuple[float, float]) -> tuple[float, float]:
    return ((p1[0] + p2[0]) / 2.0, (p1[1] + p2[1]) / 2.0)


def _draw_smooth_path(
    c: canvas_module.Canvas, points: list[tuple[float, float]], do_fill: bool, do_stroke: bool
) -> None:
    """waypoint列（`points`）を通る，各内部コーナーを丸めたオープンパスを描画する．

    各中間点（コーナー）を，隣接区間の中点同士を2次ベジェ（3次ベジェへ
    変換して近似）で結ぶことで丸める．丸めは隣接区間の長さのみに基づき
    局所的に決まるため，区間長の比（例えば横幅に対して縦が極端に長い等）
    に依らず鋭角の折れ曲がりや不自然なオーバーシュートが出ない．
    始点・終点はそのまま（丸めない）．
    """

    if not do_fill and not do_stroke:
        return
    path = c.beginPath()
    path.moveTo(*points[0])
    for i in range(1, len(points) - 1):
        corner = points[i]
        mid_in = _midpoint(points[i - 1], corner)
        mid_out = _midpoint(corner, points[i + 1])
        path.lineTo(*mid_in)
        c1 = (mid_in[0] + (corner[0] - mid_in[0]) * 2.0 / 3.0, mid_in[1] + (corner[1] - mid_in[1]) * 2.0 / 3.0)
        c2 = (mid_out[0] + (corner[0] - mid_out[0]) * 2.0 / 3.0, mid_out[1] + (corner[1] - mid_out[1]) * 2.0 / 3.0)
        path.curveTo(c1[0], c1[1], c2[0], c2[1], mid_out[0], mid_out[1])
    path.lineTo(*points[-1])
    c.drawPath(path, fill=do_fill, stroke=do_stroke)


def _draw_connector_path(
    c: canvas_module.Canvas, points: list[tuple[float, float]], is_curved: bool, do_stroke: bool
) -> None:
    """コネクタの経路（`points`の折れ線）を描画する．

    `is_curved`が真の場合は`_draw_smooth_path`で角を丸めて描画し，偽の場合
    （`bentConnector`等）は直線のまま描画する．
    """

    if not do_stroke:
        return
    if is_curved and len(points) >= 3:
        _draw_smooth_path(c, points, do_fill=False, do_stroke=True)
        return
    path = c.beginPath()
    path.moveTo(*points[0])
    for pt in points[1:]:
        path.lineTo(*pt)
    c.drawPath(path, fill=0, stroke=1)


def _draw_connector_arrowheads(
    c: canvas_module.Canvas, points: list[tuple[float, float]], is_curved: bool, stroke
) -> None:
    """コネクタの始点・終点に矢印（`headEnd`/`tailEnd`）を描画する．

    引数:
        c (canvas_module.Canvas): 描画対象のキャンバス．
        points (list[tuple[float, float]]): `_connector_points`が返す点列．
        is_curved (bool): 曲線コネクタかどうか（接線方向の計算に使用）．
        stroke: 図形の`Stroke`（矢印の色・太さ・端点種別の取得元）．
    戻り値:
        なし．
    """

    if stroke.kind != "solid" or stroke.color is None:
        return
    if stroke.head_arrow == "none" and stroke.tail_arrow == "none":
        return

    c.saveState()
    c.setFillColorRGB(*stroke.color.to_unit_tuple())
    c.setFillAlpha(stroke.alpha)
    size_pt = max(6.0, stroke.width_pt * 4.0)

    if stroke.tail_arrow != "none":
        tip = points[-1]
        away = _unit_vector(points[-2], points[-1])
        _draw_arrowhead(c, tip, away, size_pt)
    if stroke.head_arrow != "none":
        tip = points[0]
        away = _unit_vector(points[1], points[0])
        _draw_arrowhead(c, tip, away, size_pt)
    c.restoreState()


def _unit_vector(from_pt: tuple[float, float], to_pt: tuple[float, float]) -> tuple[float, float]:
    dx, dy = to_pt[0] - from_pt[0], to_pt[1] - from_pt[1]
    length = math.hypot(dx, dy) or 1.0
    return dx / length, dy / length


def _draw_arrowhead(c: canvas_module.Canvas, tip: tuple[float, float], direction: tuple[float, float], size_pt: float) -> None:
    """矢尻（三角形）を`tip`へ，`direction`方向を指すように描画する．"""

    dx, dy = direction
    px, py = -dy, dx  # directionに垂直な単位ベクトル
    half_width = size_pt * 0.35
    base_x, base_y = tip[0] - dx * size_pt, tip[1] - dy * size_pt
    p1 = (base_x + px * half_width, base_y + py * half_width)
    p2 = (base_x - px * half_width, base_y - py * half_width)
    _draw_polygon(c, [tip, p1, p2], True, False)


def draw_auto_shape(
    c: canvas_module.Canvas,
    shape: AutoShape,
    rect_pt: RectPt,
    warning_log: WarningLog,
    slide_index: int | None,
) -> None:
    """`AutoShape`の図形部分（塗り・枠線）をベクター描画する．

    引数:
        c (canvas_module.Canvas): 描画対象のキャンバス．
        shape (AutoShape): 描画対象の図形（Slide IR）．
        rect_pt (RectPt): PDF座標系での図形の矩形．
        warning_log (WarningLog): 警告記録先．
        slide_index (int | None): 対象スライド番号．
    戻り値:
        なし．
    """

    w, h = rect_pt.width, rect_pt.height
    preset = shape.preset

    c.saveState()
    apply_shape_transform(c, rect_pt, shape.rotation, shape.flip_h, shape.flip_v)
    do_fill, do_stroke = _set_fill_stroke(c, shape.style)

    if not do_fill and not do_stroke:
        c.restoreState()
        return

    if preset in ("rect", "snip1Rect", "flowChartProcess"):
        c.rect(0, 0, w, h, fill=do_fill, stroke=do_stroke)
    elif preset == "custGeom":
        if shape.custom_paths:
            _draw_custom_paths(c, shape.custom_paths, w, h, do_fill, do_stroke)
        else:
            c.rect(0, 0, w, h, fill=do_fill, stroke=do_stroke)
    elif preset in ("wedgeRectCallout", "roundRectCallout", "wedgeRoundRectCallout", "wedgeEllipseCallout", "cloudCallout"):
        _draw_polygon(c, _callout_points(w, h, shape.adjustments), do_fill, do_stroke)
    elif preset in ("roundRect", "round2SameRect", "round2DiagRect"):
        adj = shape.adjustments.get("adj", 16667) / 100000.0
        radius = min(w, h) * adj
        radius = max(0.0, min(radius, min(w, h) / 2.0))
        c.roundRect(0, 0, w, h, radius, fill=do_fill, stroke=do_stroke)
    elif preset in ("ellipse", "circle"):
        c.ellipse(0, 0, w, h, fill=do_fill, stroke=do_stroke)
    elif preset in (
        "line", "straightConnector1",
        "bentConnector2", "bentConnector3", "bentConnector4", "bentConnector5",
        "curvedConnector2", "curvedConnector3", "curvedConnector4", "curvedConnector5",
    ):
        is_curved = preset.startswith("curved")
        points = _connector_points(w, h, preset, shape.start_connect_idx, shape.end_connect_idx)
        _draw_connector_path(c, points, is_curved, do_stroke)
        _draw_connector_arrowheads(c, points, is_curved, shape.style.stroke)
    elif preset == "triangle":
        _draw_polygon(c, [(0, 0), (w, 0), (w / 2.0, h)], do_fill, do_stroke)
    elif preset in ("rightArrow", "leftArrow", "upArrow", "downArrow"):
        direction = preset.replace("Arrow", "").lower()
        _draw_polygon(c, _block_arrow_points(w, h, direction), do_fill, do_stroke)
    elif preset in ("leftRightArrow", "upDownArrow"):
        axis = "horizontal" if preset == "leftRightArrow" else "vertical"
        _draw_polygon(c, _double_block_arrow_points(w, h, axis), do_fill, do_stroke)
    elif preset in ("leftBrace", "rightBrace"):
        _draw_brace(c, w, h, do_fill, do_stroke, mirror=(preset == "leftBrace"))
    elif preset == "diamond":
        _draw_polygon(c, [(w / 2.0, 0), (w, h / 2.0), (w / 2.0, h), (0, h / 2.0)], do_fill, do_stroke)
    elif preset == "parallelogram":
        skew = w * 0.2
        _draw_polygon(c, [(skew, 0), (w, 0), (w - skew, h), (0, h)], do_fill, do_stroke)
    elif preset == "trapezoid":
        skew = w * 0.2
        _draw_polygon(c, [(skew, 0), (w - skew, 0), (w, h), (0, h)], do_fill, do_stroke)
    elif preset == "pentagon":
        _draw_polygon(c, _regular_polygon_points(w, h, 5), do_fill, do_stroke)
    elif preset == "hexagon":
        _draw_polygon(c, _regular_polygon_points(w, h, 6, start_angle_deg=0), do_fill, do_stroke)
    elif preset == "octagon":
        _draw_polygon(c, _regular_polygon_points(w, h, 8, start_angle_deg=22.5), do_fill, do_stroke)
    elif preset in ("star5", "star4", "star6"):
        n = {"star4": 4, "star5": 5, "star6": 6}[preset]
        _draw_polygon(c, _star_points(w, h, n), do_fill, do_stroke)
    elif preset == "plus":
        bar = min(w, h) * 0.25
        cx, cy = w / 2.0, h / 2.0
        _draw_polygon(
            c,
            [
                (cx - bar, 0), (cx + bar, 0), (cx + bar, cy - bar), (w, cy - bar),
                (w, cy + bar), (cx + bar, cy + bar), (cx + bar, h), (cx - bar, h),
                (cx - bar, cy + bar), (0, cy + bar), (0, cy - bar), (cx - bar, cy - bar),
            ],
            do_fill,
            do_stroke,
        )
    else:
        warning_log.unsupported_shape(preset, slide_index)
        c.rect(0, 0, w, h, fill=do_fill, stroke=do_stroke)

    c.restoreState()
