"""座標系および単位変換を一元管理するモジュール．

PPTXの座標系（EMU，左上原点，Y軸下向き）から，PDFの座標系
（pt，左下原点，Y軸上向き）への変換をここに集約する．
各要素のレンダラーは，このモジュールが提供する関数・クラスのみを
用いて座標変換を行い，個別に変換式を実装してはならない．
"""

from __future__ import annotations

from dataclasses import dataclass

EMU_PER_INCH = 914400
PT_PER_INCH = 72.0
EMU_PER_PT = EMU_PER_INCH / PT_PER_INCH  # 12700.0


def emu_to_pt(value_emu: float) -> float:
    """EMU単位の長さをpt単位へ変換する．

    引数:
        value_emu (float): EMU単位の長さ．
    戻り値:
        float: pt単位の長さ．
    """

    return value_emu / EMU_PER_PT


def pt_to_emu(value_pt: float) -> float:
    """pt単位の長さをEMU単位へ変換する．

    引数:
        value_pt (float): pt単位の長さ．
    戻り値:
        float: EMU単位の長さ．
    """

    return value_pt * EMU_PER_PT


@dataclass(frozen=True)
class RectEMU:
    """PPTX座標系（EMU，左上原点）における矩形領域．

    属性:
        x (float): 左上のX座標（EMU）．
        y (float): 左上のY座標（EMU）．
        cx (float): 幅（EMU）．
        cy (float): 高さ（EMU）．
    """

    x: float
    y: float
    cx: float
    cy: float

    @property
    def left(self) -> float:
        return self.x

    @property
    def top(self) -> float:
        return self.y

    @property
    def right(self) -> float:
        return self.x + self.cx

    @property
    def bottom(self) -> float:
        return self.y + self.cy

    @property
    def center_x(self) -> float:
        return self.x + self.cx / 2.0

    @property
    def center_y(self) -> float:
        return self.y + self.cy / 2.0


@dataclass(frozen=True)
class RectPt:
    """PDF座標系（pt，左下原点）における矩形領域．

    属性:
        x (float): 左下のX座標（pt）．
        y (float): 左下のY座標（pt）．
        width (float): 幅（pt）．
        height (float): 高さ（pt）．
    """

    x: float
    y: float
    width: float
    height: float


class CoordinateTransformer:
    """スライド1枚分のEMU座標系からPDFのpt座標系への変換を行うクラス．

    スライドサイズ（EMU）を基準に，各図形のEMU座標を正規化した上で，
    PDFページの座標系（左下原点，Y軸上向き）へ変換する．
    """

    def __init__(self, slide_width_emu: float, slide_height_emu: float) -> None:
        """コンストラクタ．

        引数:
            slide_width_emu (float): スライド幅（EMU）．
            slide_height_emu (float): スライド高さ（EMU）．
        """

        self.slide_width_emu = slide_width_emu
        self.slide_height_emu = slide_height_emu

    @property
    def page_width_pt(self) -> float:
        """PDFページ幅（pt）．"""

        return emu_to_pt(self.slide_width_emu)

    @property
    def page_height_pt(self) -> float:
        """PDFページ高さ（pt）．"""

        return emu_to_pt(self.slide_height_emu)

    def point_to_pdf(self, x_emu: float, y_emu: float) -> tuple[float, float]:
        """PPTX座標系上の1点をPDF座標系上の1点へ変換する．

        引数:
            x_emu (float): PPTX座標系のX座標（EMU，左上原点）．
            y_emu (float): PPTX座標系のY座標（EMU，左上原点）．
        戻り値:
            tuple[float, float]: PDF座標系の(x, y)（pt，左下原点）．
        """

        x_pt = emu_to_pt(x_emu)
        y_pt = self.page_height_pt - emu_to_pt(y_emu)
        return x_pt, y_pt

    def rect_to_pdf(self, rect: RectEMU) -> RectPt:
        """PPTX座標系上の矩形をPDF座標系上の矩形へ変換する．

        引数:
            rect (RectEMU): PPTX座標系の矩形（左上原点）．
        戻り値:
            RectPt: PDF座標系の矩形（左下原点で指定する矩形）．
        """

        x_pt = emu_to_pt(rect.x)
        width_pt = emu_to_pt(rect.cx)
        height_pt = emu_to_pt(rect.cy)
        y_top_pt = self.page_height_pt - emu_to_pt(rect.y)
        y_bottom_pt = y_top_pt - height_pt
        return RectPt(x=x_pt, y=y_bottom_pt, width=width_pt, height=height_pt)
