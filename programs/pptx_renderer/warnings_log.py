"""レンダリング中の警告・未対応要素を一元的に記録するモジュール．

未対応のPPTX要素やフォントの代替使用を黙って無視せず，全て記録した上で
標準エラー出力へ表示する．`--debug` 指定時には，このログの内容を
そのままデバッグ情報として出力する．
"""

from __future__ import annotations

import sys
from dataclasses import dataclass, field


@dataclass
class WarningEntry:
    """1件の警告を表すデータクラス．

    属性:
        category (str): 警告の種別（例: "unsupported_element", "font_fallback"）．
        message (str): 警告メッセージ本文．
        slide_index (int | None): 発生したスライド番号（1始まり）．未特定の場合はNone．
    """

    category: str
    message: str
    slide_index: int | None = None

    def format(self) -> str:
        """人間可読な1行の警告文字列を生成する．

        引数:
            なし．
        戻り値:
            str: フォーマット済みの警告文字列．
        """

        location = f"[slide {self.slide_index}] " if self.slide_index is not None else ""
        return f"WARNING: {location}{self.message}"


class WarningLog:
    """レンダリング中に発生した警告を蓄積するロガー．"""

    def __init__(self, echo_stderr: bool = True) -> None:
        """コンストラクタ．

        引数:
            echo_stderr (bool): 警告発生時に標準エラー出力へ即時出力するか．
        戻り値:
            なし．
        """

        self._entries: list[WarningEntry] = []
        self._echo_stderr = echo_stderr

    def add(self, category: str, message: str, slide_index: int | None = None) -> None:
        """警告を1件記録する．

        引数:
            category (str): 警告の種別．
            message (str): 警告メッセージ．
            slide_index (int | None): 対象スライド番号（1始まり）．
        戻り値:
            なし．
        """

        entry = WarningEntry(category=category, message=message, slide_index=slide_index)
        self._entries.append(entry)
        if self._echo_stderr:
            print(entry.format(), file=sys.stderr)

    def unsupported_element(self, tag: str, slide_index: int | None = None) -> None:
        """未対応のPPTX要素を記録する．

        引数:
            tag (str): 未対応であった要素のタグ名（例: "p:graphicFrame"）．
            slide_index (int | None): 対象スライド番号．
        戻り値:
            なし．
        """

        self.add("unsupported_element", f"Unsupported element: {tag}", slide_index)

    def unsupported_shape(self, prst: str, slide_index: int | None = None) -> None:
        """未対応の図形種別を記録する．

        引数:
            prst (str): 未対応であった `prstGeom` の種別名．
            slide_index (int | None): 対象スライド番号．
        戻り値:
            なし．
        """

        self.add("unsupported_shape", f"Unsupported shape type: {prst}", slide_index)

    def font_not_found(self, requested: str, fallback: str, slide_index: int | None = None) -> None:
        """フォントが見つからず代替フォントを使用したことを記録する．

        引数:
            requested (str): PPTXで指定されたフォント名．
            fallback (str): 代替として使用したフォント名．
            slide_index (int | None): 対象スライド番号．
        戻り値:
            なし．
        """

        self.add(
            "font_fallback",
            f'Font "{requested}" is not installed.\nFallback font "{fallback}" is used.',
            slide_index,
        )

    @property
    def entries(self) -> list[WarningEntry]:
        """記録済みの警告一覧．"""

        return list(self._entries)

    def summary_text(self) -> str:
        """全警告をまとめたテキストを生成する．

        引数:
            なし．
        戻り値:
            str: 警告一覧のテキスト（警告が無い場合は空文字列）．
        """

        return "\n".join(entry.format() for entry in self._entries)


# アプリケーション全体で共有するデフォルトのロガー．
default_log = WarningLog()
