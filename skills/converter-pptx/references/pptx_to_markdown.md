# pptx → Markdown 変換リファレンス

PowerPoint 資料を AI 駆動開発で扱いやすい Markdown へ変換するときの方針。
スライドの見た目やレイアウトを再現するのではなく、資料に含まれる**情報の構造と内容**を Markdown で読みやすく表現することが目的。

## 出力要件

- 出力形式は Markdown（`.md`）
- 出力先は変換元ファイルと同じディレクトリ。ファイル名は `変換元ファイル名.md`（`XXXXX_XXXXX.pptx` → `XXXXX_XXXXX.md`）
- スライド単位でセクションを分割する
- 各スライドのタイトルを見出しとして記載する
- 表形式の情報は Markdown テーブルで表現する
- 箇条書き、注記、手順、説明文は Markdown として自然に読める形に整理する
- 章・節・項目の構造が分かるように見出し（`#`, `##`, `###` …）を付ける
  - プレゼンテーション全体のタイトル（表紙スライド）を `#` 見出しとする
  - 各スライドのタイトルを `##` 見出しとする
  - スライド内の小見出しや項目は `###` 以降とする
- Markdown 見出しからリンク付き目次を作成する

## 変換手段

- `.pptx` 形式の場合: 組み込みの `pptx` スキル（`Skill(pptx)`）を使う。`python -m markitdown` または python-pptx でテキスト抽出が可能
- `.ppt` 形式（旧 PowerPoint 形式）の場合: 直接読み取れないため、後述の Windows 環境向け手順で `.pptx` に変換してから抽出する

### `.ppt` 形式の場合の推奨手順（Windows 環境）

`.ppt` は markitdown や python-pptx では直接読み取れない。PowerPoint COM オートメーション経由で `.pptx` に変換してからテキストを抽出する。

1. **PowerShell スクリプトの作成と実行**: bash からインラインで PowerShell コマンドを渡すとエスケープ問題が生じるため、必ず `.ps1` ファイルに書き出してから `-File` オプションで実行する
2. **日本語パスの解決**: 日本語パスや全角スラッシュ（／）を含むパスは `Presentations.Open()` に直接渡すとパス解決に失敗する場合がある。`Get-ChildItem -Recurse -Filter` でファイルを検索し、取得した `FullName` プロパティを使う
3. **`.pptx` 変換**: COM で開いたファイルを `SaveAs($dstFile, 24)` で `.pptx` として保存する（24 = `ppSaveAsOpenXMLPresentation`）
4. **テキスト抽出**: 変換後の `.pptx` を python-pptx で読み取る。コンソール出力の文字化けを防ぐため `sys.stdout.reconfigure(encoding='utf-8')` を設定する
5. **一時ファイルの削除**: 変換完了後、一時ファイルは必ず削除する

```powershell
# /tmp/convert_ppt.ps1 として保存して実行
$items = Get-ChildItem -Path "<検索ルートパス>" -Recurse -Filter "<ファイル名パターン>" -ErrorAction SilentlyContinue
$srcFile = $items[0].FullName
$dstFile = "<出力先>\converted.pptx"

$ppt = New-Object -ComObject PowerPoint.Application
$pres = $ppt.Presentations.Open($srcFile, $true, $false, $false)
$pres.SaveAs($dstFile, 24)  # 24 = ppSaveAsOpenXMLPresentation
$pres.Close()
$ppt.Quit()
```

```bash
powershell -ExecutionPolicy Bypass -File /tmp/convert_ppt.ps1
```

```python
# テキスト抽出コードの概要
import sys
sys.stdout.reconfigure(encoding='utf-8')
from pptx import Presentation

prs = Presentation("converted.pptx")
for i, slide in enumerate(prs.slides, 1):
    print(f"\n=== Slide {i} ===")
    if slide.has_notes_slide and slide.notes_slide.notes_text_frame:
        notes = slide.notes_slide.notes_text_frame.text.strip()
        if notes:
            print(f"[Notes]: {notes}")
    for shape in slide.shapes:
        if hasattr(shape, "text") and shape.text.strip():
            print(f"  [{shape.shape_type}, {shape.name}]: {shape.text}")
        if shape.has_table:
            table = shape.table
            for row in table.rows:
                cells = [cell.text for cell in row.cells]
                print(f"    Row: {cells}")
```

### 既知の制約・注意事項

- **pptx スキルの soffice.py は Windows 非対応**: `AF_UNIX` ソケットを使用しているため、Windows では `AttributeError: module 'socket' has no attribute 'AF_UNIX'` で失敗する。LibreOffice も環境にインストールされていない
- **markitdown の日本語文字化け**: Windows 環境ではコンソールのデフォルトエンコーディングが cp932 のため、`python -m markitdown` の出力で日本語が文字化けする。python-pptx を直接使用し、UTF-8 エンコーディングで出力する
- **ファイルパスの文字化け**: `Get-ChildItem` で取得した `FullName` はコンソール上では文字化けして見えるが、COM は正常にファイルを開ける

## 変換ルール

- スピーカーノート（ノートペイン）に意味のある内容が含まれる場合は、該当スライドセクションの末尾に「**ノート:**」として記載する
- スライドマスターやレイアウトマスターに由来するテンプレート的な文言（コピーライト、社名ロゴテキスト等）は、本文理解に不要なら出力しない
- スライド番号、ヘッダー、フッターの定型文言は出力しない
- 図表タイトル、注記、凡例など、内容理解に必要なテキストは残す
- 画像・図・模式図そのものは埋め込まず、含まれる文字情報や要点のみを必要に応じて文章化する
- SmartArt、図形グループ、フローチャートなどは、含まれるテキストと構造（順序、階層、関係性）を箇条書きや表で表現する
- グラフ・チャートは、タイトルとデータの要点をテキストまたは表で表現する
- 装飾的なアイコン、背景、色、フォントサイズ、アニメーション、トランジションなどの見た目情報は持ち込まない
- 表は列見出し・行見出しの対応関係が崩れないように Markdown テーブルへ変換する
- 同一スライド内に複数のテキストボックスがある場合は、意味が通る順序で整理する
- 資料として意味のあるラベルや注記は失わない
- 備考、注釈、注意書きなど、資料として意味のある文言は残す

## 整形方針

- スライドの視覚的レイアウトを模写するのではなく、内容として自然に読める Markdown 構造へ整理する
- 可読性を優先し、長大すぎる1つの表や不自然な改行を避ける
- 同じ内容が複数スライドで繰り返されている場合は不要な重複を避ける
- スライドごとに意味のまとまりを保つ
- 箇条書き、手順、注意事項、前提条件、制約事項などは区別できるよう整理する
- 項目名・説明・値・条件・備考の対応関係が崩れないようにする
- 日本語の自然な見出し名を用い、スライド番号やオブジェクト名を見出しにしない

## 禁止事項

- スライドの見た目やレイアウト再現を目的とした記述
- 変換ルールそのものや、PowerPoint/Markdown の操作説明の出力
- アニメーションやトランジションに関する記述
- 変換処理に関する言い訳や補足コメント
- 「以下は変換結果です」などの前置き
- 不要なコードブロックの使用
