# pptx-render

Microsoft PowerPointやLibreOfficeを一切使わず，Linux/WSL上だけで `.pptx` を `.pdf` に変換するレンダリングエンジンです．PowerPointのネイティブ数式（OMML）をLaTeX経由でベクター描画する独立した数式パイプラインを備えており，LaTeXから変換されたOMML数式を含むスライドを高い再現性でPDF化することを目標としています．

```text
PPTX ─▶ Parser ─▶ Slide IR ─▶ Layout / Shape / Image / Math Renderer ─▶ PDF
```

詳しいアーキテクチャ・実装状況は [document.md](document.md) を参照してください．

## 必要な環境

- Linux（Ubuntu 22.04で動作確認）／WSL2
- Python 3.10
- [uv](https://docs.astral.sh/uv/)（Python仮想環境管理）
- 以下の外部コマンドが `PATH` 上にあること（数式・フォント処理に使用）
  - `xelatex`（TeX Live）
  - `dvisvgm`
  - `fc-match` / `fc-list`（fontconfig）
  - `pdftoppm`（任意，PDF→PNG変換で目視確認する場合）

Ubuntuで未導入の場合は次でインストールできます．

```bash
sudo apt-get update
sudo apt-get install texlive-xetex texlive-latex-extra texlive-fonts-extra \
    dvisvgm fontconfig fonts-ipafont fonts-ipaexfont poppler-utils
```

`fonts/vendor/` に，Linux環境に標準では存在しないが実際のPPTXでよく使われるフォント（`BIZ UDGothic`／`BIZ UDPGothic`：Google Fonts配布のOFLライセンス版，`Cascadia Code`：Microsoft配布のOFLライセンス版）を同梱しているため，追加のフォントインストールは不要です．各フォントのライセンス条文は `programs/pptx_renderer/fonts/vendor/` 内の `OFL.txt` / `LICENSE-CascadiaCode.txt` を参照してください．

## セットアップ

```bash
git submodule add <このリポジトリ>
cd pptx-render

# Python仮想環境の作成（初回のみ）
uv venv .venv_render --python 3.10
uv pip install -p .venv_render -r requirements_render.txt

# 有効化
source .venv_render/bin/activate
export PYTHONPATH=programs
```

以降のコマンドは，この2行（`source` と `export PYTHONPATH`）を実行した状態のシェルで実行してください．

## 使い方

```bash
# 基本: input.pptx を output.pdf へ変換
python -m pptx_renderer.cli input.pptx -o output.pdf

# Slide IR・警告一覧などのデバッグ情報を標準エラー出力へ表示
python -m pptx_renderer.cli input.pptx -o output.pdf --debug

# 特定のスライドのみ変換（カンマ区切り，1始まり）
python -m pptx_renderer.cli input.pptx -o output.pdf --render-pages 1,3,5

# 並列ワーカー数を指定（既定はCPUコア数とページ数から自動決定，数式を多く含む場合に有効）
python -m pptx_renderer.cli input.pptx -o output.pdf --workers 4

# 追加のフォントディレクトリを検索対象にする（複数回指定可）
python -m pptx_renderer.cli input.pptx -o output.pdf --font-dir ./my_fonts

# 数式レンダリングのディスクキャッシュ（既定で有効）を使わない
python -m pptx_renderer.cli input.pptx -o output.pdf --no-math-cache
```

数式（xelatex/dvisvgm）のレンダリング結果は既定で `~/.cache/pptx_renderer/math/` にキャッシュされます．同じPPTXを繰り返し変換する場合，2回目以降は大幅に高速化されます．

未対応のPPTX要素やフォールバックフォントを使用した場合は，標準エラー出力に

```text
WARNING: Unsupported element: p:graphicFrame
WARNING:
Font "XXX" is not installed.
Fallback font "YYY" is used.
```

のように警告が出力されます．黙って無視されることはありません．

### 動作確認

任意の `.pptx` を変換して確認できます．

```bash
mkdir -p outputs
python -m pptx_renderer.cli your_slide.pptx -o outputs/your_slide_rendered.pdf

# PDF→PNG変換で目視確認する場合（poppler-utilsが必要）
pdftoppm -png -r 150 outputs/your_slide_rendered.pdf outputs/rendered
```

PowerPoint等でレンダリングした比較用PDFが手元にある場合，リポジトリ直下に `sample.pptx`／`sample.pdf` として置くと，`tests/` の回帰テストに実データでの変換確認（`test_sample_pptx_renders_all_slides`）が自動的に追加されます（無い場合はスキップされます）．

## テスト（回帰テスト）

```bash
python -m pytest tests/ -v
```

`tests/generate_fixtures.py` が空スライド・テキスト・日本語・英語・画像・図形・複数要素・表・数式・複合スライドの10種類のPPTXフィクスチャを `tests/fixtures/` に生成し，`tests/test_render.py` がそれぞれをPDF化して例外なく完了することを確認します（フィクスチャが未生成の場合は自動生成されます）．変換結果は `tests/rendered/` に出力されます．

## ディレクトリ構成

```text
programs/pptx_renderer/   # レンダラー本体（Pythonパッケージ）
  parser/                 # PPTX(XML) → Slide IR
  ir.py                   # Slide IR（中間表現）のデータクラス
  units.py                # EMU⇔pt座標変換
  layout/                 # テキストの折り返し・行送り計算
  math/                   # OMML→LaTeX→ベクターPDF（数式専用パイプライン）
  render/                 # Slide IR → PDF描画
  fonts/                  # フォント解決・同梱フォント
  cli.py                  # コマンドラインエントリポイント
tests/                    # 回帰テスト（フィクスチャ生成・pytest）
document.md               # 開発者向けの詳細ドキュメント
```

## よくあるトラブル

| 症状 | 原因・対処 |
| :--- | :--- |
| `xelatex: command not found` | TeX Liveが未導入．上記のaptコマンドを実行してください． |
| 数式が空白になる／`math_render_error`警告が出る | `--debug` を付けて実行し，`.tex`コンパイルエラーの詳細をログから確認してください（`math/latex_render.py`が一時ディレクトリで実行するため，通常は自動で片付きます）． |
| 日本語フォントが見つからないと警告される | `fc-list` でシステムのフォント一覧を確認してください．`BIZ UDGothic` は同梱済みのため警告対象外です． |
| `ModuleNotFoundError: No module named 'pptx_renderer'` | `export PYTHONPATH=programs` を実行し忘れています． |
