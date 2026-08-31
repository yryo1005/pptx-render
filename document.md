# document.md

本ファイルは，本リポジトリで実装した「Linux向けPPTX→PDFレンダラー」の構成・依存関係・実行方法をまとめたものである．セットアップ・使い方の簡潔な説明は [README.md](README.md) にもまとめている．

## 1. 全体構成

```text
PPTX (.pptx)
  ↓  parser/            … PPTX(XML)を解析
Slide IR (ir.py)         … PPTX非依存の中間表現
  ↓  layout/             … テキストの折り返し・行送り計算
  ↓  math/               … OMML→LaTeX→ベクターPDF（数式専用パイプライン）
  ↓  render/             … Slide IRをPDFへ描画
PDF (.pdf)
```

`programs/pptx_renderer/` が本体のPythonパッケージである．

```text
programs/pptx_renderer/
├── cli.py                 # コマンドラインエントリポイント（pptx-render相当）
├── ir.py                  # Slide IR（データクラス群）
├── units.py                # EMU⇔pt座標変換（一元管理）
├── warnings_log.py         # 警告ログ（未対応要素・フォント代替等）
├── parser/
│   ├── package.py          # PPTX(ZIP)パッケージ・リレーションシップ解決
│   ├── theme.py             # テーマ（配色・フォントパターン）解析
│   ├── common.py             # 座標変換・塗り/線解析等の共通処理
│   ├── placeholder.py        # プレースホルダの位置・既定書式の継承解決
│   ├── text_parser.py         # txBody→TextBody（段落・ラン・インライン数式）
│   ├── shape_parser.py         # p:sp/p:pic/p:cxnSp/p:grpSp→Slide IR
│   ├── custgeom_parser.py       # a:custGeom(カスタム図形)→ベクターパス
│   ├── table_parser.py          # p:graphicFrame(表)→TableShape（`ppt/tableStyles.xml`の縞模様・見出し行の配色も解決）
│   ├── slide_parser.py           # 1スライド分の解析（背景・図形ツリー）
│   └── omml_to_latex.py（math/）  # ※実体はmath/以下
├── fonts/
│   ├── registry.py          # フォント解決・ReportLab登録・メトリクス取得
│   └── vendor/               # Linuxに存在しない既知フォントを同梱
│       └── BIZUDGothic-*.ttf   # Google Fonts配布のBIZ UDGothic（OFLライセンス）
├── layout/
│   └── text_layout.py        # 折り返し（CJK文字単位／英単語単位）・行送り計算
├── math/
│   ├── omml_to_latex.py       # OMML→LaTeX変換
│   └── latex_render.py         # LaTeX→ベクターPDF（xelatex+dvisvgm+cairosvg）
└── render/
    ├── pdf_renderer.py         # 全体のオーケストレーション（グループ展開含む）
    ├── shape_renderer.py        # 既定図形のベクター描画
    ├── text_renderer.py          # レイアウト済みテキストの描画
    ├── image_renderer.py          # 画像（PNG/JPEG/SVG）描画
    ├── table_renderer.py           # 表の描画（オートフィット付き）
    └── overlay.py                   # 数式・SVGのベクターPDF合成（pypdf）
```

## 2. 依存関係

### 2.1 Python環境

`uv` で `.venv_render` を作成し，`requirements_render.txt` に基づいて依存ライブラリをインストールする．

```bash
uv venv .venv_render --python 3.10
uv pip install -p .venv_render -r requirements_render.txt
source .venv_render/bin/activate
```

主要ライブラリ: `python-pptx`（フィクスチャ生成のみで使用；PPTX解析自体は自前実装），`lxml`，`reportlab`，`pypdf`，`fonttools`，`Pillow`，`cairosvg`，`pytest`．

### 2.2 外部コマンド（システムパッケージ）

| コマンド | 用途 | 想定パッケージ（Ubuntu） |
| :--- | :--- | :--- |
| `xelatex` | OMML由来のLaTeX数式のタイプセット | `texlive-xetex` |
| `dvisvgm` | LaTeXの中間形式(XDV)からSVGへの変換（ベースライン情報取得） | `texlive-extra-utils` または `dvisvgm` |
| `fc-match` / `fc-list` | フォント解決（フォールバック検出） | `fontconfig` |
| （IPA/IPAexフォント） | 日本語フォールバックフォント | `fonts-ipafont`, `fonts-ipaexfont` |
| `pdftoppm` | 回帰確認用（PDF→PNG，本体には非依存） | `poppler-utils` |

もしこれらが存在しない環境で本レンダラーを利用する場合は，以下を実行する．

```bash
sudo apt-get install texlive-xetex texlive-latex-extra texlive-fonts-extra \
    dvisvgm fontconfig fonts-ipafont fonts-ipaexfont poppler-utils
```

### 2.3 同梱フォント（`fonts/vendor/`）

PowerPointの多くのテーマ・テンプレートは日本語フォントとして `BIZ UDGothic` を指定するが，このフォントはLinux環境に標準では存在しない．黙って別フォントへ置換するのではなく，Google Fonts（`google/fonts`リポジトリ，SIL Open Font License）で配布されている本物の `BIZ UDGothic-Regular.ttf` / `BIZUDGothic-Bold.ttf` を `programs/pptx_renderer/fonts/vendor/` に同梱し，最優先で使用する．ライセンス条文は同ディレクトリの `OFL.txt` を参照．

同様に，実データで使用が確認された `BIZ UDPGothic`（プロポーショナル版，同じくGoogle Fonts配布のOFLライセンス版）と `Cascadia Code`（Microsoft配布，SIL Open Font License）についても，本物のフォントファイルを同梱している．

| フォントファミリ名 | ファイル | ライセンス条文 |
| :--- | :--- | :--- |
| `BIZ UDGothic` | `BIZUDGothic-Regular.ttf` / `BIZUDGothic-Bold.ttf` | `OFL.txt` |
| `BIZ UDPGothic` | `BIZUDPGothic-Regular.ttf` / `BIZUDPGothic-Bold.ttf` | `OFL.txt`（`BIZ UDGothic`と著作権表示が共通のため同一ファイルで兼用） |
| `Cascadia Code` | `CascadiaCode-Regular/Bold/Italic/BoldItalic.ttf` | `LICENSE-CascadiaCode.txt`（Microsoft Corporation，SIL OFL 1.1） |

それ以外のフォントは `fc-match` によりシステムフォントから解決し，指定フォントが見つからない場合は必ず

```text
WARNING:
Font "XXX" is not installed.
Fallback font "YYY" is used.
```

の形式で警告する．

## 3. 実行方法

```bash
source .venv_render/bin/activate
export PYTHONPATH=programs   # または `uv run --project .` 等でパッケージを解決

# 基本
python -m pptx_renderer.cli input.pptx -o output.pdf

# デバッグ情報（Slide IR・警告一覧）を表示
python -m pptx_renderer.cli input.pptx -o output.pdf --debug

# 特定スライドのみ描画
python -m pptx_renderer.cli input.pptx -o output.pdf --render-pages 1,2,3

# 並列ワーカー数を指定（既定はCPUコア数とページ数から自動決定）
python -m pptx_renderer.cli input.pptx -o output.pdf --workers 4

# 追加のフォント検索ディレクトリを指定（複数回指定可）
python -m pptx_renderer.cli input.pptx -o output.pdf --font-dir ./my_fonts

# 数式レンダリングのディスクキャッシュを使用しない
python -m pptx_renderer.cli input.pptx -o output.pdf --no-math-cache
```

数式レンダリング結果は既定で `~/.cache/pptx_renderer/math/`（`$XDG_CACHE_HOME`があればそちら，`$PPTX_RENDERER_MATH_CACHE_DIR`で明示指定も可）にキャッシュされ，同一PPTXの再変換を大幅に高速化する．

## 4. テスト

```bash
source .venv_render/bin/activate
export PYTHONPATH=programs
python -m pytest tests/ -v
```

- `tests/generate_fixtures.py`: 回帰テスト用PPTXフィクスチャ（`tests/fixtures/01〜10_*.pptx`）を`python-pptx`で生成する．
- `tests/test_units.py`: 座標変換（EMU⇔pt）の単体テスト．
- `tests/test_render.py`: 10種類のフィクスチャを実際にPDF化し，例外なく完了すること・妥当なPDFであることを確認する．リポジトリ直下に `sample.pptx` を置いた場合は，それを対象とした追加テストも実行される（無い場合は自動的にスキップ）．
- `tests/rendered/`: テスト実行時に生成されるPDF（`.gitignore`対象）．

## 5. Git管理上の注意事項

- `.venv_render/`, `tests/rendered/`, `outputs/`, `__pycache__/`, `.pytest_cache/` は `.gitignore` により除外している．
- `programs/pptx_renderer/fonts/vendor/*.ttf` はサイズが大きい（同梱フォント合計 約9MB）が，OFLライセンスのフォント資産であり，レンダラーの動作に必須のためリポジトリに含める．

## 6. 既知の制約

- `a:custGeom`（カスタム図形パス）は矩形として近似する．
- `mc:AlternateContent` は `Choice` 側（SVG等の新機能）を優先的に採用する．
- 回転・反転した図形内の数式・SVG画像の合成は未対応．
- プレースホルダの既定書式（色・太字・サイズ）は，スライドレイアウト／マスターの`lstStyle`・`txStyles`から解決するが，深いネストの継承規則を完全には実装していない．
- `leftBrace`/`rightBrace`（波括弧）は，OOXMLの調整値（`adj1`＝丸みの深さ，`adj2`＝先端の垂直位置）を反映せず，固定比率のベジェ曲線で近似する．
- `p:graphicFrame`（`uri=http://schemas.openxmlformats.org/drawingml/2006/diagram`，SmartArt）は未対応．データモデル（`dgm:dataModel`）とレイアウトアルゴリズムの解決が必要な大規模機能のため，現状は「未対応の要素」として警告を出した上でスキップする．
- `bentConnector`/`curvedConnector`は，接続点番号（`idx`）が`0`〜`3`（上・左・下・右）の範囲外，またはローカルボックスの短辺が極端に小さい（回転済みの細長いコネクタ等）場合，屈曲点の計算を行わず始点・終点を直接結ぶ経路にフォールバックする．
- テーブルスタイル（`ppt/tableStyles.xml`）は，セルの塗り（`wholeTbl`/`band1H`/`band2H`/`firstRow`/`lastRow`の`fill`）のみ反映する．罫線（`tcBdr`）や見出し行の文字装飾（太字・文字色等，`tcTxStyle`）は未対応．
