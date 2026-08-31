"""PPTXで指定されたフォント名を実際のフォントファイルへ解決するモジュール．

フォントの存在確認,フォントファイルの探索,ReportLabへの登録,
フォントメトリクス取得を一元的に扱う．

解決の優先順位は以下の通りとする．

1. `fonts/vendor/` に同梱した既知フォント（PPTX生成基盤で標準的に
   使用されるが，Linux環境に標準では存在しないフォント）．
2. `_KNOWN_ALIASES` に登録された，よく使われるが本環境には存在しない
   Windows/Office標準フォント（Meiryo, Yu Gothic, MS Gothic等）を，
   日本語グリフを持つ代替フォントへ明示的にマッピングする表．
   `fontconfig`はこれらの名前に対して日本語グリフを持たないフォント
   （DejaVu Sans等）を返すことがあり，その場合日本語が空白として
   消失するため，明示的な対応表で防ぐ．
3. `fontconfig`（`fc-match`）によるシステムフォントの解決．

PPTXで指定されたフォントが見つからない場合は，代替フォントを使用した上で
必ず警告を記録する．警告を出さずに黙って代替フォントへ切り替えることは
禁止する．

なお，`fonts/vendor/`・`_KNOWN_ALIASES`のいずれにも該当しない未知の
フォント名については，本モジュールはネットワークからの自動取得を行わない
（ライセンス上再配布できないフォント，あるいは取得元が不明なフォントを
実行時に無条件でダウンロードすることを避けるため）．そのようなフォントを
高精度に再現する必要がある場合は，実際に使用するフォントファイルを人手で
`fonts/vendor/`へ追加し，本ファイルの対応表を更新することを想定する．
"""

from __future__ import annotations

import subprocess
from dataclasses import dataclass
from pathlib import Path

from fontTools.ttLib import TTFont as FontToolsTTFont
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont as ReportLabTTFont

from pptx_renderer.warnings_log import WarningLog, default_log

_VENDOR_DIR = Path(__file__).parent / "vendor"

# Linux標準では提供されないが，`ai-pptx-kit`のテンプレート・テーマ，
# および利用者が頻繁に使用する既知のPPTXフォーマット（`sample2.pptx`等）で
# 使用されるフォントを同梱し，優先的に使用する．
# キーは正規化（小文字化）したフォントファミリ名．
_VENDOR_FONTS: dict[str, dict[str, Path]] = {
    "biz udgothic": {
        "regular": _VENDOR_DIR / "BIZUDGothic-Regular.ttf",
        "bold": _VENDOR_DIR / "BIZUDGothic-Bold.ttf",
    },
    "biz udゴシック": {
        "regular": _VENDOR_DIR / "BIZUDGothic-Regular.ttf",
        "bold": _VENDOR_DIR / "BIZUDGothic-Bold.ttf",
    },
    "biz udpgothic": {
        "regular": _VENDOR_DIR / "BIZUDPGothic-Regular.ttf",
        "bold": _VENDOR_DIR / "BIZUDPGothic-Bold.ttf",
    },
    "biz udpゴシック": {
        "regular": _VENDOR_DIR / "BIZUDPGothic-Regular.ttf",
        "bold": _VENDOR_DIR / "BIZUDPGothic-Bold.ttf",
    },
    "cascadia code": {
        "regular": _VENDOR_DIR / "CascadiaCode-Regular.ttf",
        "bold": _VENDOR_DIR / "CascadiaCode-Bold.ttf",
        "italic": _VENDOR_DIR / "CascadiaCode-Italic.ttf",
        "bold_italic": _VENDOR_DIR / "CascadiaCode-BoldItalic.ttf",
    },
}

# Windows/Office標準の日本語フォントで,Linuxに同名フォントが存在しない場合に
# `fontconfig`が日本語グリフを持たないフォントへ誤って解決してしまうものを,
# 日本語グリフを持つ代替フォント名へ明示的にマッピングする．
# ここでの代替先はシステムに標準搭載されているフォント（IPA/IPAex系列）に
# 限定しており，新たなダウンロードは発生しない．
_KNOWN_JAPANESE_FONT_ALIASES: dict[str, str] = {
    "meiryo": "IPAexGothic",
    "meiryo ui": "IPAexGothic",
    "yu gothic": "IPAexGothic",
    "yu gothic ui": "IPAexGothic",
    "游ゴシック": "IPAexGothic",
    "ms gothic": "IPAGothic",
    "ｍｓ ゴシック": "IPAGothic",
    "ms pgothic": "IPAPGothic",
    "ｍｓ ｐゴシック": "IPAPGothic",
    "yu mincho": "IPAexMincho",
    "游明朝": "IPAexMincho",
    "ms mincho": "IPAMincho",
    "ｍｓ 明朝": "IPAMincho",
    "ms pmincho": "IPAPMincho",
    "ｍｓ ｐ明朝": "IPAPMincho",
    "hiragino kaku gothic pro": "IPAexGothic",
    "hiragino sans": "IPAexGothic",
    "hiragino mincho pro": "IPAexMincho",
}


@dataclass
class ResolvedFont:
    """フォント解決結果を保持するデータクラス．

    属性:
        family_requested (str): PPTXで指定されたフォント名．
        family_actual (str): 実際に使用するフォントのファミリ名．
        file_path (Path): 実際に使用するフォントファイルのパス．
        bold_requested (bool): 太字が要求されたか．
        italic_requested (bool): 斜体が要求されたか．
        bold_available (bool): 実際に使用するフォントが太字書体を持つか．
        italic_available (bool): 実際に使用するフォントが斜体書体を持つか．
        is_fallback (bool): 要求されたフォントと異なるフォントへ代替されたか．
    """

    family_requested: str
    family_actual: str
    file_path: Path
    bold_requested: bool
    italic_requested: bool
    bold_available: bool
    italic_available: bool
    is_fallback: bool

    @property
    def needs_faux_bold(self) -> bool:
        return self.bold_requested and not self.bold_available

    @property
    def needs_faux_italic(self) -> bool:
        return self.italic_requested and not self.italic_available

    @property
    def reportlab_font_name(self) -> str:
        """ReportLabへ登録する際に使用する一意なフォント名．"""

        return f"{self.family_actual}-{self.file_path.stem}".replace(" ", "_")


class FontRegistry:
    """フォント解決・登録・メトリクス取得を行うレジストリ．"""

    def __init__(
        self,
        warning_log: WarningLog | None = None,
        extra_font_dirs: list[str | Path] | None = None,
    ) -> None:
        """コンストラクタ．

        引数:
            warning_log (WarningLog | None): 警告記録先．Noneの場合は共有ロガーを使用する．
            extra_font_dirs (list[str | Path] | None): 追加のフォント検索ディレクトリ
                （CLIの`--font-dir`に対応）．`.ttf`/`.otf`を再帰的に走査し，
                同梱フォント（`fonts/vendor/`）と同様に最優先で使用する．
        戻り値:
            なし．
        """

        self._warning_log = warning_log or default_log
        self._resolve_cache: dict[tuple[str, bool, bool], ResolvedFont] = {}
        self._registered_names: set[str] = set()
        self._ttfont_cache: dict[Path, FontToolsTTFont] = {}
        self._warned_families: set[str] = set()
        self._extra_fonts: dict[str, dict[str, Path]] = {}
        for font_dir in extra_font_dirs or []:
            self._scan_font_dir(Path(font_dir))

    def _scan_font_dir(self, font_dir: Path) -> None:
        if not font_dir.is_dir():
            self._warning_log.add("font_dir_not_found", f"フォントディレクトリが見つかりません: {font_dir}", None)
            return

        for font_path in list(font_dir.rglob("*.ttf")) + list(font_dir.rglob("*.otf")):
            try:
                tt = FontToolsTTFont(str(font_path), fontNumber=0, lazy=True)
                name_table = tt["name"]
                family = name_table.getDebugName(1)
                subfamily = (name_table.getDebugName(2) or "Regular").lower()
            except Exception:  # noqa: BLE001
                continue
            if not family:
                continue

            key = family.strip().lower()
            slot = self._extra_fonts.setdefault(key, {})
            is_bold = "bold" in subfamily
            is_italic = "italic" in subfamily or "oblique" in subfamily
            if is_bold and is_italic:
                slot["bold_italic"] = font_path
            elif is_bold:
                slot["bold"] = font_path
            elif is_italic:
                slot["italic"] = font_path
            else:
                slot.setdefault("regular", font_path)

    def resolve(
        self,
        family: str,
        bold: bool,
        italic: bool,
        slide_index: int | None = None,
    ) -> ResolvedFont:
        """フォントファミリ名・太字・斜体の指定から実際のフォントを解決する．

        引数:
            family (str): PPTXで指定されたフォントファミリ名．
            bold (bool): 太字が要求されているか．
            italic (bool): 斜体が要求されているか．
            slide_index (int | None): 対象スライド番号（警告表示用）．
        戻り値:
            ResolvedFont: 解決結果．
        """

        cache_key = (family, bold, italic)
        if cache_key in self._resolve_cache:
            return self._resolve_cache[cache_key]

        normalized = family.strip().lower()
        extra = self._extra_fonts.get(normalized)
        vendor = _VENDOR_FONTS.get(normalized)
        alias_target = _KNOWN_JAPANESE_FONT_ALIASES.get(normalized)
        if extra is not None:
            result = self._resolve_from_local_dict(family, extra, bold, italic)
        elif vendor is not None:
            result = self._resolve_from_local_dict(family, vendor, bold, italic)
        elif alias_target is not None:
            result = self._resolve_via_fontconfig(alias_target, bold, italic, slide_index, requested_as=family)
        else:
            result = self._resolve_via_fontconfig(family, bold, italic, slide_index)

        self._resolve_cache[cache_key] = result
        return result

    def _resolve_from_local_dict(
        self, family: str, styles: dict[str, Path], bold: bool, italic: bool
    ) -> ResolvedFont:
        """同梱フォント（`fonts/vendor/`）または`--font-dir`で追加されたフォントを解決する．"""

        if bold and italic and "bold_italic" in styles:
            path, bold_available, italic_available = styles["bold_italic"], True, True
        elif bold and "bold" in styles:
            path, bold_available, italic_available = styles["bold"], True, False
        elif italic and "italic" in styles:
            path, bold_available, italic_available = styles["italic"], False, True
        else:
            path, bold_available, italic_available = styles.get("regular"), False, False

        return ResolvedFont(
            family_requested=family,
            family_actual=family,
            file_path=path,
            bold_requested=bold,
            italic_requested=italic,
            bold_available=bold_available,
            italic_available=italic_available,
            is_fallback=False,
        )

    def _resolve_via_fontconfig(
        self,
        family: str,
        bold: bool,
        italic: bool,
        slide_index: int | None,
        requested_as: str | None = None,
    ) -> ResolvedFont:
        """`fc-match`によりフォントを解決する．

        引数:
            family (str): `fc-match`への問い合わせに使用するフォント名．
            bold (bool): 太字が要求されているか．
            italic (bool): 斜体が要求されているか．
            slide_index (int | None): 対象スライド番号．
            requested_as (str | None): PPTX上で実際に指定されていたフォント名．
                `_KNOWN_JAPANESE_FONT_ALIASES`経由で別名へ読み替えて問い合わせる
                場合，警告文言や`ResolvedFont.family_requested`にはこちらを表示する．
                Noneの場合は`family`をそのまま用いる．
        戻り値:
            ResolvedFont: 解決結果．
        """

        display_name = requested_as if requested_as is not None else family
        style_tokens = []
        if bold:
            style_tokens.append("bold")
        if italic:
            style_tokens.append("italic")
        pattern = family if not style_tokens else f"{family}:{':'.join(style_tokens)}"

        try:
            proc = subprocess.run(
                ["fc-match", "-f", "%{file}\t%{family}\t%{style}\n", pattern],
                capture_output=True,
                text=True,
                timeout=10,
            )
            file_str, matched_family, style = proc.stdout.strip().split("\t")
        except Exception:  # noqa: BLE001
            file_str, matched_family, style = "", "", ""

        if not file_str:
            proc = subprocess.run(
                ["fc-match", "-f", "%{file}\t%{family}\t%{style}\n", "sans-serif"],
                capture_output=True,
                text=True,
                timeout=10,
            )
            file_str, matched_family, style = proc.stdout.strip().split("\t")
            self._warn_fallback(display_name, matched_family.split(",")[0].strip(), slide_index)
            return ResolvedFont(
                family_requested=display_name,
                family_actual=matched_family.split(",")[0].strip(),
                file_path=Path(file_str),
                bold_requested=bold,
                italic_requested=italic,
                bold_available="bold" in style.lower(),
                italic_available="italic" in style.lower(),
                is_fallback=True,
            )

        matched_primary_family = matched_family.split(",")[0].strip()
        is_fallback = requested_as is not None or matched_primary_family.lower() != family.strip().lower()
        if is_fallback:
            self._warn_fallback(display_name, matched_primary_family, slide_index)

        return ResolvedFont(
            family_requested=display_name,
            family_actual=matched_primary_family,
            file_path=Path(file_str),
            bold_requested=bold,
            italic_requested=italic,
            bold_available="bold" in style.lower(),
            italic_available="italic" in style.lower() or "oblique" in style.lower(),
            is_fallback=is_fallback,
        )

    def _warn_fallback(self, requested: str, fallback: str, slide_index: int | None) -> None:
        key = requested.strip().lower()
        if key in self._warned_families:
            return
        self._warned_families.add(key)
        self._warning_log.font_not_found(requested, fallback, slide_index)

    def register_reportlab_font(self, resolved: ResolvedFont) -> str:
        """解決済みフォントをReportLabへ登録し，登録名を返す．

        引数:
            resolved (ResolvedFont): `resolve()`の戻り値．
        戻り値:
            str: `reportlab.pdfbase.pdfmetrics`へ登録済みのフォント名．
        """

        name = resolved.reportlab_font_name
        if name not in self._registered_names:
            pdfmetrics.registerFont(ReportLabTTFont(name, str(resolved.file_path)))
            self._registered_names.add(name)
        return name

    def get_ttfont(self, resolved: ResolvedFont) -> FontToolsTTFont:
        """メトリクス取得用に`fontTools`の`TTFont`オブジェクトを取得する．

        引数:
            resolved (ResolvedFont): `resolve()`の戻り値．
        戻り値:
            fontTools.ttLib.TTFont: フォントメトリクス参照用オブジェクト．
        """

        if resolved.file_path not in self._ttfont_cache:
            self._ttfont_cache[resolved.file_path] = FontToolsTTFont(
                str(resolved.file_path), fontNumber=0, lazy=True
            )
        return self._ttfont_cache[resolved.file_path]


# アプリケーション全体で共有するデフォルトのフォントレジストリ．
default_registry = FontRegistry()
