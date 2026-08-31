"""LaTeX数式文字列をベクターPDFへレンダリングするモジュール．

処理系統は独立したパイプラインとして構成する．

```text
LaTeX数式文字列
 ↓ xelatex -no-pdf
XDVファイル
 ↓ dvisvgm --bbox=preview --no-fonts
SVGファイル（文字はパス化済み，ベースライン情報を stderr から取得）
 ↓ cairosvg
ベクターPDF（1オブジェクト分の小さいPDF）
```

`preview`パッケージのtightpage機能と`dvisvgm`の`--bbox=preview`オプションを
組み合わせることで，数式のベースライン位置（ascent/descentに相当する
height/depth）をpt単位で正確に取得できる．これにより，本文テキストの
ベースラインと数式のベースラインを一致させて合成できる．

日本語フォントは数式そのものには通常使用しないため，このモジュールは
数式用OpenType数学フォント（Latin Modern Math）に限定して使用する．
"""

from __future__ import annotations

import hashlib
import os
import re
import subprocess
import tempfile
from dataclasses import dataclass
from pathlib import Path

import cairosvg

from pptx_renderer.warnings_log import WarningLog, default_log

_MATH_FONT = "texgyretermes-math.otf"

# 数式レンダリング結果の永続キャッシュ先．同一のPPTXを何度も変換する
# 反復作業（スライドを少し直しては再変換する等）で，同じ数式を毎回
# xelatex/dvisvgmで再コンパイルすることを避けるために使用する．
# `_MATH_FONT`を含めてキーを作るため，数式フォントを変更した場合は
# 自動的にキャッシュが無効化される．
_DISK_CACHE_DIR = Path(
    os.environ.get("PPTX_RENDERER_MATH_CACHE_DIR")
    or (Path(os.environ.get("XDG_CACHE_HOME", Path.home() / ".cache")) / "pptx_renderer" / "math")
)

_TEMPLATE = r"""\documentclass[border=0pt]{{standalone}}
\usepackage[active,tightpage,dvips]{{preview}}
\usepackage{{amsmath,amssymb,mathtools,bm,xcolor}}
\usepackage{{unicode-math}}
\setmathfont{{{math_font}}}
\begin{{document}}
\begin{{preview}}
\fontsize{{{size_pt}}}{{{leading_pt}}}\selectfont
\mbox{{${display}{body}$}}
\end{{preview}}
\end{{document}}
"""

_DVISVGM_RE = re.compile(
    r"width=(?P<width>[-\d.]+)pt,\s*height=(?P<height>[-\d.]+)pt,\s*depth=(?P<depth>[-\d.]+)pt"
)


@dataclass
class MathRenderResult:
    """レンダリング済み数式の情報を保持するデータクラス．

    属性:
        pdf_bytes (bytes): 数式1個分のベクターPDFのバイト列．
        width_pt (float): 数式の幅（pt）．
        height_pt (float): ベースラインより上側の高さ（pt）．
        depth_pt (float): ベースラインより下側の深さ（pt）．
    """

    pdf_bytes: bytes
    width_pt: float
    height_pt: float
    depth_pt: float

    @property
    def total_height_pt(self) -> float:
        return self.height_pt + self.depth_pt


class LatexMathRenderer:
    """LaTeX数式文字列からベクターPDFを生成し，結果をキャッシュするクラス．"""

    def __init__(self, warning_log: WarningLog | None = None, use_disk_cache: bool = True) -> None:
        """コンストラクタ．

        引数:
            warning_log (WarningLog | None): 警告記録先．Noneの場合は共有ロガーを使用する．
            use_disk_cache (bool): ディスクキャッシュ（`~/.cache/pptx_renderer/math/`等）を
                使用するか．同一の数式を含むPPTXを繰り返し変換する場合の高速化に有効．
        戻り値:
            なし．
        """

        self._cache: dict[str, MathRenderResult | None] = {}
        self._warning_log = warning_log or default_log
        self._use_disk_cache = use_disk_cache

    def render(
        self,
        latex_body: str,
        size_pt: float,
        display: bool = False,
        slide_index: int | None = None,
    ) -> MathRenderResult | None:
        """LaTeX数式本体をレンダリングし，結果を返す．

        引数:
            latex_body (str): 数式本体のLaTeXコード（`$`を含まない）．
            size_pt (float): 数式のフォントサイズ（pt）．
            display (bool): ディスプレイスタイル（`\\displaystyle`）で組むかどうか．
            slide_index (int | None): 対象スライド番号（警告表示用）．
        戻り値:
            MathRenderResult | None: レンダリング結果．失敗した場合はNone．
        """

        cache_key = hashlib.sha256(
            f"{_MATH_FONT}|{size_pt}|{display}|{latex_body}".encode("utf-8")
        ).hexdigest()
        if cache_key in self._cache:
            return self._cache[cache_key]

        if self._use_disk_cache:
            cached = self._load_from_disk(cache_key)
            if cached is not None:
                self._cache[cache_key] = cached
                return cached

        result = self._render_uncached(latex_body, size_pt, display, slide_index)
        self._cache[cache_key] = result
        if self._use_disk_cache and result is not None:
            self._save_to_disk(cache_key, result)
        return result

    def _load_from_disk(self, cache_key: str) -> MathRenderResult | None:
        pdf_path = _DISK_CACHE_DIR / f"{cache_key}.pdf"
        dims_path = _DISK_CACHE_DIR / f"{cache_key}.dims"
        if not pdf_path.exists() or not dims_path.exists():
            return None
        try:
            width_pt, height_pt, depth_pt = (float(v) for v in dims_path.read_text().split())
            return MathRenderResult(
                pdf_bytes=pdf_path.read_bytes(), width_pt=width_pt, height_pt=height_pt, depth_pt=depth_pt
            )
        except (ValueError, OSError):
            return None

    def _save_to_disk(self, cache_key: str, result: MathRenderResult) -> None:
        # 並列レンダリング時に複数スレッドが同じキーへ同時に書き込んでも
        # 読み取り側が破損した内容を見ないよう，一時ファイルに書いてから
        # 目的のパスへ原子的にリネームする．
        try:
            _DISK_CACHE_DIR.mkdir(parents=True, exist_ok=True)
            pdf_path = _DISK_CACHE_DIR / f"{cache_key}.pdf"
            dims_path = _DISK_CACHE_DIR / f"{cache_key}.dims"
            tmp_pdf = pdf_path.with_suffix(f".pdf.tmp{os.getpid()}{id(result)}")
            tmp_dims = dims_path.with_suffix(f".dims.tmp{os.getpid()}{id(result)}")
            tmp_pdf.write_bytes(result.pdf_bytes)
            tmp_dims.write_text(f"{result.width_pt} {result.height_pt} {result.depth_pt}")
            os.replace(tmp_pdf, pdf_path)
            os.replace(tmp_dims, dims_path)
        except OSError:
            pass  # キャッシュ書き込みの失敗はレンダリング自体には影響させない．

    def _render_uncached(
        self,
        latex_body: str,
        size_pt: float,
        display: bool,
        slide_index: int | None,
    ) -> MathRenderResult | None:
        leading_pt = size_pt * 1.2
        source = _TEMPLATE.format(
            math_font=_MATH_FONT,
            size_pt=f"{size_pt:.2f}",
            leading_pt=f"{leading_pt:.2f}",
            display=r"\displaystyle " if display else "",
            body=latex_body,
        )

        with tempfile.TemporaryDirectory(prefix="pptx_render_math_") as tmpdir:
            work_dir = Path(tmpdir)
            tex_path = work_dir / "eq.tex"
            tex_path.write_text(source, encoding="utf-8")

            xelatex_ok = self._run_xelatex(tex_path, work_dir)
            if not xelatex_ok:
                self._warning_log.add(
                    "math_render_error",
                    f"数式のLaTeXコンパイルに失敗しました．数式: {latex_body!r}",
                    slide_index,
                )
                return None

            xdv_path = work_dir / "eq.xdv"
            svg_path = work_dir / "eq.svg"
            dims = self._run_dvisvgm(xdv_path, svg_path)
            if dims is None:
                self._warning_log.add(
                    "math_render_error",
                    f"数式のSVG変換に失敗しました．数式: {latex_body!r}",
                    slide_index,
                )
                return None
            width_pt, height_pt, depth_pt = dims

            pdf_path = work_dir / "eq.pdf"
            try:
                cairosvg.svg2pdf(url=str(svg_path), write_to=str(pdf_path))
            except Exception as exc:  # noqa: BLE001 - 外部ライブラリの例外を包括的に捕捉
                self._warning_log.add(
                    "math_render_error",
                    f"数式PDFの生成に失敗しました: {exc}",
                    slide_index,
                )
                return None

            pdf_bytes = pdf_path.read_bytes()

        return MathRenderResult(
            pdf_bytes=pdf_bytes,
            width_pt=width_pt,
            height_pt=height_pt,
            depth_pt=depth_pt,
        )

    @staticmethod
    def _run_xelatex(tex_path: Path, work_dir: Path) -> bool:
        try:
            proc = subprocess.run(
                [
                    "xelatex",
                    "-no-pdf",
                    "-interaction=nonstopmode",
                    "-halt-on-error",
                    tex_path.name,
                ],
                cwd=work_dir,
                capture_output=True,
                timeout=30,
            )
        except (subprocess.TimeoutExpired, FileNotFoundError):
            return False
        return proc.returncode == 0 and (work_dir / "eq.xdv").exists()

    @staticmethod
    def _run_dvisvgm(xdv_path: Path, svg_path: Path) -> tuple[float, float, float] | None:
        try:
            proc = subprocess.run(
                [
                    "dvisvgm",
                    "--no-fonts",
                    "--bbox=preview",
                    "-o",
                    str(svg_path),
                    str(xdv_path),
                ],
                capture_output=True,
                text=True,
                timeout=30,
            )
        except (subprocess.TimeoutExpired, FileNotFoundError):
            return None
        if proc.returncode != 0 or not svg_path.exists():
            return None

        match = _DVISVGM_RE.search(proc.stderr)
        if not match:
            return None
        return (
            float(match.group("width")),
            float(match.group("height")),
            float(match.group("depth")),
        )
