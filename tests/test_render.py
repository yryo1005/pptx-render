"""回帰テスト．

`tests/fixtures/*.pptx`（1.空スライド〜10.複合スライド）を入力として，
`PPTX -> Slide IR -> PDF`の変換が例外なく完了し，妥当なPDFが生成される
ことを確認する．リポジトリルートに`sample.pptx`（実データ）を置いた場合は，
それを入力とした変換テストも追加で実行される（無い場合は自動的にスキップ）．

画素単位の一致確認（PowerPointとの見た目の完全一致）は行わない．
"""

from __future__ import annotations

from pathlib import Path

import pytest
from pypdf import PdfReader

from pptx_renderer.parser import parse_presentation
from pptx_renderer.render.pdf_renderer import PdfRenderer
from pptx_renderer.warnings_log import WarningLog

FIXTURES_DIR = Path(__file__).parent / "fixtures"
RENDERED_DIR = Path(__file__).parent / "rendered"
REPO_ROOT = Path(__file__).parent.parent

FIXTURE_NAMES = [
    "01_empty_slide",
    "02_text",
    "03_japanese",
    "04_english",
    "05_image",
    "06_shapes",
    "07_multi",
    "08_table",
    "09_equation",
    "10_complex",
]


@pytest.fixture(scope="session", autouse=True)
def _ensure_fixtures_exist():
    if not FIXTURES_DIR.exists() or not any(FIXTURES_DIR.glob("*.pptx")):
        import subprocess
        import sys

        subprocess.run([sys.executable, str(Path(__file__).parent / "generate_fixtures.py")], check=True)
    RENDERED_DIR.mkdir(parents=True, exist_ok=True)


@pytest.mark.parametrize("name", FIXTURE_NAMES)
def test_fixture_renders_without_error(name: str) -> None:
    """各フィクスチャがエラーなくPDF化され，妥当なPDFファイルになることを確認する．"""

    pptx_path = FIXTURES_DIR / f"{name}.pptx"
    output_path = RENDERED_DIR / f"{name}.pdf"

    warning_log = WarningLog(echo_stderr=False)
    presentation = parse_presentation(str(pptx_path), warning_log)

    renderer = PdfRenderer(warning_log=warning_log)
    renderer.render(presentation, str(output_path))

    assert output_path.exists()
    assert output_path.stat().st_size > 0

    reader = PdfReader(str(output_path))
    assert len(reader.pages) == len(presentation.slides)
    assert len(reader.pages) >= 1


def test_empty_slide_has_no_shapes() -> None:
    presentation = parse_presentation(str(FIXTURES_DIR / "01_empty_slide.pptx"))
    assert len(presentation.slides) == 1
    assert presentation.slides[0].shapes == []


def test_table_fixture_has_merged_cell() -> None:
    presentation = parse_presentation(str(FIXTURES_DIR / "08_table.pptx"))
    from pptx_renderer.ir import TableShape

    table = next(s for s in presentation.slides[0].shapes if isinstance(s, TableShape))
    covered_cells = [c for row in table.rows for c in row.cells if c.is_covered]
    assert len(covered_cells) >= 1


def test_equation_fixture_produces_math_run() -> None:
    from pptx_renderer.ir import MathRun

    presentation = parse_presentation(str(FIXTURES_DIR / "09_equation.pptx"))
    shape = presentation.slides[0].shapes[0]
    math_runs = [
        run
        for paragraph in shape.text_body.paragraphs
        for run in paragraph.runs
        if isinstance(run, MathRun)
    ]
    assert len(math_runs) == 1
    assert r"\frac" in math_runs[0].latex_body


@pytest.mark.skipif(not (REPO_ROOT / "sample.pptx").exists(), reason="sample.pptx が存在しません")
def test_sample_pptx_renders_all_slides() -> None:
    """数式レンダリングを含む実データ（sample.pptx）が
    全ページ例外なくレンダリングできることを確認する．
    """

    sample_path = REPO_ROOT / "sample.pptx"
    output_path = RENDERED_DIR / "sample.pdf"

    warning_log = WarningLog(echo_stderr=False)
    presentation = parse_presentation(str(sample_path), warning_log)

    renderer = PdfRenderer(warning_log=warning_log)
    renderer.render(presentation, str(output_path))

    reader = PdfReader(str(output_path))
    assert len(reader.pages) == len(presentation.slides)

    error_categories = {"math_render_error", "image_render_error"}
    errors = [e for e in warning_log.entries if e.category in error_categories]
    assert errors == [], f"レンダリングエラーが発生しました: {[e.format() for e in errors]}"
