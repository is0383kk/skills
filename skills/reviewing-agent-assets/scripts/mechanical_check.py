#!/usr/bin/env python3
"""AIエージェント向け資産の品質レビュー用・機械的チェックスクリプト。

決定論的に確認できる項目（行数・目次要否の目安・バックスラッシュのパス表記・
数値リテラル・用語ゆれ）を出力し、AIの判定の根拠として使う。
標準ライブラリのみで動作する（追加インストール不要）。

使い方:
    python scripts/mechanical_check.py <対象パス> [--terms 語1,語2,...]

    <対象パス>: ファイルまたはディレクトリ。ディレクトリは .md/.txt を再帰的に対象にする。
    --terms:    全文検索して混在を確認したい語をカンマ区切りで指定（用語統一の確認用）。
"""

from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

TARGET_SUFFIXES = {".md", ".txt"}
LINE_THRESHOLD = 100  # 100行を超えるファイルは目次要否を確認する（対象チェックリストの基準）
NUMERIC_PATTERN = re.compile(r"\b\d+\b")


def collect_files(target: Path) -> list[Path]:
    """対象ファイル一覧を返す。ファイルならそれ自身、ディレクトリなら .md/.txt を再帰収集。"""
    if target.is_file():
        return [target]
    return sorted(
        p for p in target.rglob("*")
        if p.is_file() and p.suffix.lower() in TARGET_SUFFIXES
    )


def read_lines(path: Path) -> list[str]:
    """UTF-8 で読み込み、行リストを返す（改行は含めない）。"""
    text = path.read_text(encoding="utf-8", errors="replace")
    return text.splitlines()


def rel_path(path: Path, root: Path) -> Path:
    """root からの相対パスを返す（辿れない場合は元のパスのまま）。"""
    try:
        return path.relative_to(root)
    except ValueError:
        return path


def check_file(path: Path, root: Path) -> dict:
    """1ファイルの機械的チェック結果を返す。"""
    lines = read_lines(path)
    backslash_lines = [
        (i + 1, ln) for i, ln in enumerate(lines) if "\\" in ln
    ]
    numeric_lines = [
        (i + 1, ln) for i, ln in enumerate(lines)
        if NUMERIC_PATTERN.search(ln)
    ]
    return {
        "path": rel_path(path, root),
        "line_count": len(lines),
        "over_threshold": len(lines) > LINE_THRESHOLD,
        "backslash_lines": backslash_lines,
        "numeric_lines": numeric_lines,
    }


def count_terms(
    files: list[Path], terms: list[str]
) -> dict[str, list[tuple[Path, int]]]:
    """語ごとに、出現したファイルと件数を返す（用語ゆれ確認用）。"""
    result: dict[str, list[tuple[Path, int]]] = {t: [] for t in terms}
    for f in files:
        text = f.read_text(encoding="utf-8", errors="replace")
        for t in terms:
            n = text.count(t)
            if n > 0:
                result[t].append((f, n))
    return result


def main(argv: list[str]) -> int:
    # Windows コンソール等の非 UTF-8 環境でも文字化け・クラッシュしないよう UTF-8 出力に統一する
    for stream in (sys.stdout, sys.stderr):
        reconfigure = getattr(stream, "reconfigure", None)
        if reconfigure is not None:
            reconfigure(encoding="utf-8")

    parser = argparse.ArgumentParser(description="AIエージェント向け資産の機械的チェック")
    parser.add_argument("target", help="対象ファイルまたはディレクトリ")
    parser.add_argument("--terms", default="", help="全文検索する語（カンマ区切り）")
    args = parser.parse_args(argv)

    target = Path(args.target)
    if not target.exists():
        print(f"[エラー] 対象が存在しません: {target}", file=sys.stderr)
        return 2

    files = collect_files(target)
    if not files:
        print(f"[エラー] 対象に .md/.txt が見つかりません: {target}", file=sys.stderr)
        return 2

    root = target if target.is_dir() else target.parent

    print("=" * 60)
    print(f"対象: {target}  （{len(files)} ファイル）")
    print("=" * 60)

    for f in files:
        r = check_file(f, root)
        print(f"\n## {r['path']}")
        flag = "  ← 100行超: 目次の有無を確認" if r["over_threshold"] else ""
        print(f"- 行数: {r['line_count']}{flag}")

        if r["backslash_lines"]:
            print(f"- バックスラッシュ(\\)を含む行: {len(r['backslash_lines'])} 件")
            for ln_no, ln in r["backslash_lines"]:
                print(f"    {ln_no}: {ln.strip()}")
        else:
            print("- バックスラッシュ(\\)を含む行: なし")

        if r["numeric_lines"]:
            print(f"- 数値リテラルを含む行: {len(r['numeric_lines'])} 件（根拠の有無を確認）")
            for ln_no, ln in r["numeric_lines"]:
                print(f"    {ln_no}: {ln.strip()}")
        else:
            print("- 数値リテラルを含む行: なし")

    terms = [t.strip() for t in args.terms.split(",") if t.strip()]
    if terms:
        print("\n" + "=" * 60)
        print(f"用語ゆれチェック: {', '.join(terms)}")
        print("=" * 60)
        counts = count_terms(files, terms)
        for t in terms:
            hits = counts[t]
            total = sum(n for _, n in hits)
            print(f"\n- 「{t}」: 合計 {total} 件")
            for f, n in hits:
                print(f"    {rel_path(f, root)}: {n} 件")
        used = [t for t in terms if sum(n for _, n in counts[t]) > 0]
        if len(used) >= 2:
            print(
                f"\n[注意] {len(used)} 語が混在しています"
                f"（{', '.join(used)}）。用語統一を検討してください。"
            )

    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
