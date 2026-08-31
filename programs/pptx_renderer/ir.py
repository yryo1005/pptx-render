"""PPTXの内容を表す中間表現（Slide IR）のデータモデル．

PPTXのXML構造とPDFの描画処理を分離するため，パーサーはPPTXのXMLを直接
描画せず，必ずこのモジュールが定義するデータクラス群（Slide IR）へ変換する．
レンダラーは，このSlide IRのみを入力として受け取り，PPTXのXML構造を
直接参照してはならない．
"""

from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum

from pptx_renderer.units import RectEMU


class HAlign(Enum):
    """段落の水平方向の配置．"""

    LEFT = "left"
    CENTER = "center"
    RIGHT = "right"
    JUSTIFY = "justify"


class VAlign(Enum):
    """テキストボックス内の垂直方向の配置．"""

    TOP = "top"
    MIDDLE = "middle"
    BOTTOM = "bottom"


@dataclass(frozen=True)
class RGBColor:
    """RGB形式の色を表すデータクラス．

    属性:
        r (int): 赤成分（0〜255）．
        g (int): 緑成分（0〜255）．
        b (int): 青成分（0〜255）．
    """

    r: int
    g: int
    b: int

    @classmethod
    def from_hex(cls, hex_str: str) -> "RGBColor":
        """16進数文字列（例: "FF0000"）からRGBColorを生成する．

        引数:
            hex_str (str): 6桁の16進数カラーコード．
        戻り値:
            RGBColor: 変換後の色．
        """

        hex_str = hex_str.lstrip("#")
        return cls(
            r=int(hex_str[0:2], 16),
            g=int(hex_str[2:4], 16),
            b=int(hex_str[4:6], 16),
        )

    def to_hex(self) -> str:
        """16進数文字列表現を返す．

        引数:
            なし．
        戻り値:
            str: 6桁の16進数カラーコード（先頭に"#"は付けない）．
        """

        return f"{self.r:02X}{self.g:02X}{self.b:02X}"

    def to_unit_tuple(self) -> tuple[float, float, float]:
        """0〜1に正規化したRGBタプルを返す（ReportLab用）．

        引数:
            なし．
        戻り値:
            tuple[float, float, float]: (r, g, b)の各成分（0〜1）．
        """

        return (self.r / 255.0, self.g / 255.0, self.b / 255.0)


BLACK = RGBColor(0, 0, 0)
WHITE = RGBColor(255, 255, 255)


@dataclass
class Fill:
    """図形・背景の塗りつぶし情報．

    属性:
        kind (str): "none" または "solid"．
        color (RGBColor | None): 塗りつぶし色．kindが"solid"の場合に使用する．
        alpha (float): 不透明度（0.0〜1.0，1.0が不透明）．
    """

    kind: str = "none"
    color: RGBColor | None = None
    alpha: float = 1.0


@dataclass
class Stroke:
    """図形の枠線情報．

    属性:
        kind (str): "none" または "solid"．
        color (RGBColor | None): 枠線色．
        width_pt (float): 線幅（pt）．
        alpha (float): 不透明度（0.0〜1.0）．
        dash (str): 破線スタイル（"solid", "dash", "dashDot", "sysDot" 等）．
    """

    kind: str = "none"
    color: RGBColor | None = None
    width_pt: float = 0.75
    alpha: float = 1.0
    dash: str = "solid"
    head_arrow: str = "none"
    tail_arrow: str = "none"


@dataclass
class TextRun:
    """通常のテキストラン（1つの書式が適用される文字列の単位）．

    属性:
        text (str): 表示文字列．
        font_name (str | None): PPTXで指定されたフォント名（未指定の場合None）．
        size_pt (float): フォントサイズ（pt）．
        bold (bool): 太字か．
        italic (bool): 斜体か．
        underline (bool): 下線を引くか．
        color (RGBColor): 文字色．
        alpha (float): 不透明度．
    """

    text: str
    font_name: str | None
    size_pt: float
    ea_font_name: str | None = None
    char_spacing_pt: float = 0.0
    bold: bool = False
    italic: bool = False
    underline: bool = False
    color: RGBColor = field(default_factory=lambda: BLACK)
    alpha: float = 1.0


@dataclass
class LineBreak:
    """明示的な改行（`a:br`）を表すラン．"""


@dataclass
class MathRun:
    """数式ラン（OMMLから変換されたLaTeX数式）．

    属性:
        latex_body (str): 数式本体のLaTeXコード（`$`を含まない）．
        size_pt (float): 数式のフォントサイズ（pt）．
        display (bool): 段落全体が数式のみで構成される（`m:oMathPara`）場合はTrue．
    """

    latex_body: str
    size_pt: float
    display: bool = False


RunIR = TextRun | MathRun | LineBreak


@dataclass
class Paragraph:
    """段落．複数のランから構成される．

    属性:
        runs (list[RunIR]): 段落内のラン列．
        align (HAlign): 水平方向の配置．
        line_spacing_pct (float | None): 行間（%指定，100が1行分）．
        line_spacing_pt (float | None): 行間（pt指定）．指定時はpct指定より優先する．
        space_before_pt (float): 段落前の空白（pt）．
        space_after_pt (float): 段落後の空白（pt）．
        level (int): 箇条書きのインデントレベル．
    """

    runs: list[RunIR] = field(default_factory=list)
    align: HAlign = HAlign.LEFT
    line_spacing_pct: float | None = None
    line_spacing_pt: float | None = None
    space_before_pt: float = 0.0
    space_after_pt: float = 0.0
    level: int = 0
    indent_left_emu: float = 0.0
    empty_line_size_pt: float | None = None
    bullet_char: str | None = None
    bullet_font: str | None = None
    bullet_offset_emu: float = 0.0
    bullet_auto_num_fmt: str | None = None


@dataclass
class TextBody:
    """テキストボックス・図形内のテキスト本体．

    属性:
        paragraphs (list[Paragraph]): 段落のリスト．
        anchor (VAlign): 垂直方向の配置．
        wrap (bool): テキストを矩形幅で折り返すか．
        inset_left_emu (float): 左内部余白（EMU）．
        inset_top_emu (float): 上内部余白（EMU）．
        inset_right_emu (float): 右内部余白（EMU）．
        inset_bottom_emu (float): 下内部余白（EMU）．
        font_scale (float): オートフィットによるフォント縮小率（1.0が等倍）．
    """

    paragraphs: list[Paragraph] = field(default_factory=list)
    anchor: VAlign = VAlign.TOP
    wrap: bool = True
    inset_left_emu: float = 91440.0
    inset_top_emu: float = 45720.0
    inset_right_emu: float = 91440.0
    inset_bottom_emu: float = 45720.0
    font_scale: float = 1.0


@dataclass
class ShapeStyle:
    """図形の共通スタイル情報（塗り・枠線）．"""

    fill: Fill = field(default_factory=Fill)
    stroke: Stroke = field(default_factory=Stroke)


@dataclass
class AutoShape:
    """既定図形およびテキストボックス（`p:sp`）．

    PPTXのテキストボックスは内部的に`prstGeom prst="rect"`の図形として
    表現されるため，テキストボックスと既定図形（四角形・円・線・矢印等）を
    区別せず，同一のデータクラスで表現する．
    """

    id: str
    name: str
    rect: RectEMU
    preset: str
    style: ShapeStyle = field(default_factory=ShapeStyle)
    text_body: TextBody | None = None
    rotation: float = 0.0
    flip_h: bool = False
    flip_v: bool = False
    adjustments: dict[str, float] = field(default_factory=dict)
    custom_paths: list["CustomPath"] = field(default_factory=list)
    start_connect_idx: int | None = None
    end_connect_idx: int | None = None


@dataclass
class CustomPath:
    """`a:custGeom`から抽出した1つのパス（`a:path`）のベクター描画コマンド列．

    属性:
        width_emu (float): パスのローカル座標系における幅（`a:path/@w`）．
        height_emu (float): パスのローカル座標系における高さ（`a:path/@h`）．
        commands (list[tuple]): 描画コマンドのリスト．各要素は
            ("moveTo", x, y), ("lineTo", x, y),
            ("curveTo", x1, y1, x2, y2, x3, y3), ("close",) のいずれか
            （座標は全てEMU相当，パスのローカル座標系，Y軸下向き）．
        fill (bool): この個別パスを塗りつぶし対象に含めるか．
        stroke (bool): この個別パスを線描画対象に含めるか．
    """

    width_emu: float
    height_emu: float
    commands: list[tuple] = field(default_factory=list)
    fill: bool = True
    stroke: bool = True


@dataclass
class PictureShape:
    """画像（`p:pic`）．"""

    id: str
    name: str
    rect: RectEMU
    image_bytes: bytes
    image_format: str
    crop_left: float = 0.0
    crop_top: float = 0.0
    crop_right: float = 0.0
    crop_bottom: float = 0.0
    alpha: float = 1.0
    style: ShapeStyle = field(default_factory=ShapeStyle)
    rotation: float = 0.0
    flip_h: bool = False
    flip_v: bool = False


@dataclass
class TableCell:
    """表のセル．"""

    text_body: TextBody | None
    fill: Fill = field(default_factory=Fill)
    col_span: int = 1
    row_span: int = 1
    is_covered: bool = False
    border_left: Stroke = field(default_factory=Stroke)
    border_right: Stroke = field(default_factory=Stroke)
    border_top: Stroke = field(default_factory=Stroke)
    border_bottom: Stroke = field(default_factory=Stroke)


@dataclass
class TableRow:
    """表の行．"""

    height_emu: float
    cells: list[TableCell] = field(default_factory=list)


@dataclass
class TableShape:
    """表（`p:graphicFrame`のうち`a:tbl`を含むもの）．"""

    id: str
    name: str
    rect: RectEMU
    col_widths_emu: list[float]
    rows: list[TableRow] = field(default_factory=list)
    rotation: float = 0.0


@dataclass
class GroupShape:
    """グループ化された図形（`p:grpSp`）．"""

    id: str
    name: str
    rect: RectEMU
    children: list["ShapeIR"] = field(default_factory=list)
    child_offset_emu: tuple[float, float] = (0.0, 0.0)
    child_extent_emu: tuple[float, float] = (1.0, 1.0)
    rotation: float = 0.0
    flip_h: bool = False
    flip_v: bool = False


ShapeIR = AutoShape | PictureShape | TableShape | GroupShape


@dataclass
class Background:
    """スライド背景．"""

    fill: Fill = field(default_factory=lambda: Fill(kind="solid", color=WHITE))


@dataclass
class Slide:
    """1枚のスライドを表すSlide IRのルート．

    属性:
        index (int): スライド番号（1始まり）．
        width_emu (float): スライド幅（EMU）．
        height_emu (float): スライド高さ（EMU）．
        background (Background): 背景情報．
        shapes (list[ShapeIR]): 描画順（背面から前面）に並んだ図形のリスト．
    """

    index: int
    width_emu: float
    height_emu: float
    background: Background = field(default_factory=Background)
    shapes: list[ShapeIR] = field(default_factory=list)


@dataclass
class Presentation:
    """PPTXファイル全体を表すSlide IRのルート．

    属性:
        width_emu (float): スライド幅（EMU）．
        height_emu (float): スライド高さ（EMU）．
        slides (list[Slide]): スライドのリスト．
    """

    width_emu: float
    height_emu: float
    slides: list[Slide] = field(default_factory=list)
