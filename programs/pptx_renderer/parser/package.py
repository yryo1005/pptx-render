"""PPTX（OOXML）パッケージ（ZIP）へのアクセスを提供するモジュール．

PPTXファイル内の各パート（XMLファイルやメディアファイル）の読み込み，
リレーションシップ（`_rels`）の解決，スライド一覧の取得等，パッケージ構造に
関する処理をここに集約する．
"""

from __future__ import annotations

import posixpath
import zipfile
from dataclasses import dataclass

from lxml import etree

_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
_R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
_A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"


@dataclass
class Relationship:
    """1件のリレーションシップ情報．

    属性:
        rid (str): リレーションシップID（例: "rId1"）．
        rel_type (str): リレーションシップの種別URI．
        target (str): 解決済みのターゲットパート名（ZIP内の絶対パス）．
    """

    rid: str
    rel_type: str
    target: str


class PptxPackage:
    """PPTXファイル（ZIPパッケージ）へのアクセスを提供するクラス．"""

    def __init__(self, pptx_path: str) -> None:
        """コンストラクタ．

        引数:
            pptx_path (str): PPTXファイルのパス．
        戻り値:
            なし．
        """

        self._zip = zipfile.ZipFile(pptx_path)
        self._xml_cache: dict[str, etree._Element] = {}
        self._rels_cache: dict[str, dict[str, Relationship]] = {}

    def read_bytes(self, part_name: str) -> bytes:
        """パートの生バイト列を取得する．

        引数:
            part_name (str): ZIP内のパート名（例: "ppt/media/image1.png"）．
        戻り値:
            bytes: パートの内容．
        """

        return self._zip.read(part_name)

    def read_xml(self, part_name: str) -> etree._Element:
        """パートをXML要素として取得する（キャッシュ付き）．

        引数:
            part_name (str): ZIP内のパート名．
        戻り値:
            etree._Element: パースされたXMLのルート要素．
        """

        if part_name not in self._xml_cache:
            self._xml_cache[part_name] = etree.fromstring(self.read_bytes(part_name))
        return self._xml_cache[part_name]

    def exists(self, part_name: str) -> bool:
        """指定パートがパッケージ内に存在するかを確認する．"""

        return part_name in self._zip.namelist()

    def relationships_for(self, part_name: str) -> dict[str, Relationship]:
        """指定パートに対応するリレーションシップ一覧を取得する．

        引数:
            part_name (str): 対象パート名（例: "ppt/slides/slide1.xml"）．
        戻り値:
            dict[str, Relationship]: rIdをキーとするリレーションシップの辞書．
        """

        if part_name in self._rels_cache:
            return self._rels_cache[part_name]

        directory, filename = posixpath.split(part_name)
        rels_part = posixpath.join(directory, "_rels", f"{filename}.rels")

        result: dict[str, Relationship] = {}
        if self.exists(rels_part):
            root = self.read_xml(rels_part)
            for rel_el in root.findall(f"{{{_REL_NS}}}Relationship"):
                rid = rel_el.get("Id")
                rel_type = rel_el.get("Type")
                target = rel_el.get("Target")
                mode = rel_el.get("TargetMode", "Internal")
                if mode == "External":
                    resolved = target
                else:
                    resolved = posixpath.normpath(posixpath.join(directory, target))
                result[rid] = Relationship(rid=rid, rel_type=rel_type, target=resolved)

        self._rels_cache[part_name] = result
        return result

    def resolve_rid(self, part_name: str, rid: str) -> str | None:
        """指定パート内で使われる`r:id`から解決済みターゲットパート名を取得する．

        引数:
            part_name (str): rIdを参照している側のパート名．
            rid (str): リレーションシップID．
        戻り値:
            str | None: 解決済みターゲットパート名．見つからない場合はNone．
        """

        rel = self.relationships_for(part_name).get(rid)
        return rel.target if rel else None

    def slide_part_names(self) -> list[str]:
        """`presentation.xml`の`p:sldIdLst`の順序でスライドのパート名一覧を取得する．

        引数:
            なし．
        戻り値:
            list[str]: スライドのパート名（例: "ppt/slides/slide1.xml"）のリスト．
        """

        pres_part = "ppt/presentation.xml"
        root = self.read_xml(pres_part)
        sld_id_lst = root.find(f"{{{_P_NS}}}sldIdLst")
        part_names = []
        if sld_id_lst is not None:
            for sld_id in sld_id_lst.findall(f"{{{_P_NS}}}sldId"):
                rid = sld_id.get(f"{{{_R_NS}}}id")
                target = self.resolve_rid(pres_part, rid)
                if target:
                    part_names.append(target)
        return part_names

    def slide_size_emu(self) -> tuple[float, float]:
        """スライドサイズ（EMU）を取得する．

        引数:
            なし．
        戻り値:
            tuple[float, float]: (幅, 高さ)のEMU値．
        """

        root = self.read_xml("ppt/presentation.xml")
        sld_sz = root.find(f"{{{_P_NS}}}sldSz")
        return float(sld_sz.get("cx")), float(sld_sz.get("cy"))

    def slide_layout_for_slide(self, slide_part: str) -> str | None:
        """スライドに対応するスライドレイアウトのパート名を取得する．"""

        for rel in self.relationships_for(slide_part).values():
            if rel.rel_type.endswith("/slideLayout"):
                return rel.target
        return None

    def slide_master_for_layout(self, layout_part: str) -> str | None:
        """スライドレイアウトに対応するスライドマスターのパート名を取得する．"""

        for rel in self.relationships_for(layout_part).values():
            if rel.rel_type.endswith("/slideMaster"):
                return rel.target
        return None

    def theme_for_master(self, master_part: str) -> str | None:
        """スライドマスターに対応するテーマのパート名を取得する．"""

        for rel in self.relationships_for(master_part).values():
            if rel.rel_type.endswith("/theme"):
                return rel.target
        return None

    def theme_for_slide(self, slide_part: str) -> str | None:
        """スライドから辿ってテーマのパート名を取得する（layout→master→theme）．

        引数:
            slide_part (str): スライドのパート名．
        戻り値:
            str | None: テーマのパート名．解決できない場合はNone．
        """

        layout_part = self.slide_layout_for_slide(slide_part)
        if layout_part is None:
            return None
        master_part = self.slide_master_for_layout(layout_part)
        if master_part is None:
            return None
        return self.theme_for_master(master_part)
