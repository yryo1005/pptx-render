"""PPTX全体をSlide IR（`pptx_renderer.ir.Presentation`）へ変換するエントリポイント．"""

from __future__ import annotations

from pptx_renderer.ir import Presentation
from pptx_renderer.parser.package import PptxPackage
from pptx_renderer.parser.slide_parser import parse_slide
from pptx_renderer.warnings_log import WarningLog, default_log


def parse_presentation(pptx_path: str, warning_log: WarningLog | None = None) -> Presentation:
    """PPTXファイルを解析し，`Presentation`（Slide IR）を構築する．

    引数:
        pptx_path (str): PPTXファイルのパス．
        warning_log (WarningLog | None): 警告記録先．Noneの場合は共有ロガーを使用する．
    戻り値:
        Presentation: 全スライドを含む中間表現．
    """

    log = warning_log or default_log
    package = PptxPackage(pptx_path)
    width_emu, height_emu = package.slide_size_emu()

    slides = []
    for index, slide_part in enumerate(package.slide_part_names(), start=1):
        slides.append(parse_slide(package, slide_part, index, width_emu, height_emu, log))

    return Presentation(width_emu=width_emu, height_emu=height_emu, slides=slides)
