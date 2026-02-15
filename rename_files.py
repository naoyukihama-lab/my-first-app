#!/usr/bin/env python3
"""
rename_files.py - ファイル内容からタイトルを抽出して自動リネームするスクリプト

使い方:
  python rename_files.py <ファイルまたはディレクトリ> [オプション]

例:
  python rename_files.py ./documents/
  python rename_files.py report.pdf memo.docx
  python rename_files.py ./docs/ --recursive --yes

依存ライブラリ（オプション）:
  pip install pypdf          # PDF テキスト抽出を強化（なくても動作）
"""

import argparse
import json
import os
import re
import sys
import zipfile
from pathlib import Path


# ---------------------------------------------------------------------------
# ファイル名から日本語タイトルを生成
# ---------------------------------------------------------------------------

def clean_filename(name: str) -> str:
    """ファイル名（拡張子なし）を日本語タイトル形式に変換する"""
    # YYYYMMDDHHMMSS（14桁）
    m = re.match(r'^(\d{4})(\d{2})(\d{2})(\d{2})(\d{2})(\d{2})$', name)
    if m:
        return f"{m[1]}年{m[2]}月{m[3]}日 {m[4]}:{m[5]}:{m[6]}"

    # YYYYMMDDHHMM（12桁）
    m = re.match(r'^(\d{4})(\d{2})(\d{2})(\d{2})(\d{2})$', name)
    if m:
        return f"{m[1]}年{m[2]}月{m[3]}日 {m[4]}:{m[5]}"

    # YYYYMMDD（8桁）
    m = re.match(r'^(\d{4})(\d{2})(\d{2})$', name)
    if m:
        return f"{m[1]}年{m[2]}月{m[3]}日"

    # YYYYMM + テキスト（例: 202206kohan）
    m = re.match(r'^(\d{4})(\d{2})([^\d].*)$', name)
    if m:
        rest = clean_filename(re.sub(r'^[-_\s]+', '', m.group(3)))
        return f"{m.group(1)}年{m.group(2)}月 {rest}"

    # 区切り文字をスペースに展開・camelCase を分割
    result = re.sub(r'[-_\.]+', ' ', name)
    result = re.sub(r'([a-z])([A-Z])', r'\1 \2', result)
    result = re.sub(r'([A-Z]+)([A-Z][a-z])', r'\1 \2', result)
    result = ' '.join(w.capitalize() for w in result.split() if w)
    return result or name


# ---------------------------------------------------------------------------
# テキストコンテンツからタイトル抽出
# ---------------------------------------------------------------------------

def extract_title_from_text(content: str, ext: str) -> str | None:
    lines = content.split('\n')

    # Markdown / RST: 最初の # 見出し
    if ext in ('md', 'markdown', 'rst'):
        for line in lines:
            m = re.match(r'^#{1,3}\s+(.+)', line)
            if m:
                return re.sub(r'\s*#{1,3}\s*$', '', m.group(1)).strip()

    # HTML: <title> または <h1>
    if ext in ('html', 'htm'):
        m = re.search(r'<title[^>]*>([^<]+)</title>', content, re.IGNORECASE)
        if m:
            return m.group(1).strip()
        m = re.search(r'<h1[^>]*>([^<]+)</h1>', content, re.IGNORECASE)
        if m:
            return re.sub(r'<[^>]+>', '', m.group(1)).strip()

    # JSON: title / name / description フィールド
    if ext == 'json':
        try:
            obj = json.loads(content)
            for key in ('title', 'name', 'description', 'label'):
                val = obj.get(key)
                if val and isinstance(val, str):
                    return val.strip()[:80]
        except Exception:
            pass

    # SVG: <title>
    if ext == 'svg':
        m = re.search(r'<title[^>]*>([^<]+)</title>', content, re.IGNORECASE)
        if m:
            return m.group(1).strip()

    # CSV / TSV: 1行目をヘッダーとして解釈
    if ext in ('csv', 'tsv'):
        sep = '\t' if ext == 'tsv' else ','
        cols = [c.strip().strip('"\'') for c in lines[0].split(sep)]
        if cols and cols[0]:
            return ' / '.join(cols)[:60]

    # YAML / TOML: title: / name: キー
    if ext in ('yaml', 'yml', 'toml'):
        for line in lines:
            m = re.match(r'^(?:title|name)\s*[:=]\s*["\']?(.+?)["\']?\s*$', line, re.IGNORECASE)
            if m:
                return m.group(1).strip()

    # コードファイル: 先頭コメント → クラス名 → 関数名
    code_exts = {
        'js', 'mjs', 'ts', 'tsx', 'jsx', 'vue', 'svelte',
        'py', 'rb', 'php', 'java', 'kt', 'swift', 'go', 'rs',
        'cpp', 'c', 'cs', 'sh', 'bash', 'zsh', 'ps1', 'sql', 'r', 'lua',
        'css', 'scss', 'sass',
    }
    if ext in code_exts:
        for line in lines[:20]:
            s = line.strip()
            m = re.match(
                r'^(?://|#|--|%|;)\s*@?(?:title|name|brief|desc(?:ription)?)[:：\s]+(.+)',
                s, re.IGNORECASE,
            )
            if m:
                return m.group(1).strip()
            m = re.match(r'^(?://|#|--)\s+([A-Za-z\u3040-\u9FFF].{4,60})$', s)
            if m and not re.match(r'^[*\-=]+$', m.group(1)):
                return m.group(1).strip()
        for line in lines[:50]:
            m = re.search(r'(?:class|struct|interface|object)\s+([A-Z]\w+)', line)
            if m:
                return re.sub(r'([A-Z])', r' \1', m.group(1)).strip()
            m = re.search(r'(?:def|func|function|fn|sub)\s+([a-z_]\w+)', line, re.IGNORECASE)
            if m and m.group(1).lower() != 'main':
                return clean_filename(m.group(1))

    # 汎用: 最初の意味ある行
    for line in lines[:10]:
        text = re.sub(r'^[#*\-=_/\\|<>!@`~]+\s*', '', line).strip()
        if 4 <= len(text) <= 100:
            return text[:80]

    return None


# ---------------------------------------------------------------------------
# PDF からタイトル抽出
# ---------------------------------------------------------------------------

def extract_title_from_pdf(path: Path) -> str | None:
    # pypdf が入っていれば優先使用
    try:
        import pypdf  # type: ignore
        reader = pypdf.PdfReader(str(path))
        meta = reader.metadata
        if meta and getattr(meta, 'title', None) and meta.title.strip():
            return meta.title.strip()[:80]
        if reader.pages:
            text = reader.pages[0].extract_text() or ''
            for line in text.split('\n'):
                line = line.strip()
                if len(line) >= 4:
                    return line[:80]
    except ImportError:
        pass
    except Exception:
        pass

    # フォールバック: バイナリから /Title エントリを探す
    try:
        with open(path, 'rb') as f:
            raw = f.read(65536)
        text = raw.decode('latin-1', errors='ignore')
        m = re.search(r'/Title\s*\(([^)]{2,80})\)', text)
        if m:
            title = m.group(1).strip()
            if title:
                return title[:80]
    except Exception:
        pass

    return None


# ---------------------------------------------------------------------------
# Office Open XML（DOCX / XLSX / PPTX）からタイトル抽出
# ---------------------------------------------------------------------------

def extract_title_from_office(path: Path, ext: str) -> str | None:
    try:
        with zipfile.ZipFile(path, 'r') as z:
            names = z.namelist()

            # docProps/core.xml の dc:title を最優先
            if 'docProps/core.xml' in names:
                xml = z.read('docProps/core.xml').decode('utf-8', errors='ignore')
                m = re.search(r'<dc:title>([^<]+)</dc:title>', xml, re.IGNORECASE)
                if m and m.group(1).strip():
                    return m.group(1).strip()[:80]

            if ext == 'docx' and 'word/document.xml' in names:
                xml = z.read('word/document.xml').decode('utf-8', errors='ignore')
                for para in re.findall(r'<w:p[ >].*?</w:p>', xml, re.DOTALL):
                    parts = re.findall(r'<w:t[^>]*>([^<]+)</w:t>', para)
                    text = ''.join(parts).strip()
                    if len(text) >= 4:
                        return text[:80]

            if ext == 'xlsx':
                wb = next((n for n in names if re.match(r'xl/workbook\.xml', n)), None)
                if wb:
                    xml = z.read(wb).decode('utf-8', errors='ignore')
                    m = re.search(r'<sheet[^>]+name="([^"]+)"', xml, re.IGNORECASE)
                    if m and m.group(1).strip():
                        return m.group(1).strip()

            if ext == 'pptx':
                slides = sorted(n for n in names if re.match(r'ppt/slides/slide\d+\.xml', n))
                if slides:
                    xml = z.read(slides[0]).decode('utf-8', errors='ignore')
                    m = re.search(r'<a:t>([^<]{2,})</a:t>', xml)
                    if m and m.group(1).strip():
                        return m.group(1).strip()[:80]
    except Exception:
        pass

    return None


# ---------------------------------------------------------------------------
# タイトル生成メイン
# ---------------------------------------------------------------------------

TEXT_EXTS = {
    'txt', 'md', 'markdown', 'rst',
    'html', 'htm', 'css', 'scss', 'sass',
    'json', 'xml', 'csv', 'tsv', 'yaml', 'yml', 'toml',
    'js', 'mjs', 'ts', 'tsx', 'jsx', 'vue', 'svelte',
    'py', 'rb', 'php', 'java', 'kt', 'swift', 'go', 'rs',
    'cpp', 'c', 'cs', 'sh', 'bash', 'zsh', 'ps1', 'sql', 'r', 'lua', 'svg',
}
OFFICE_EXTS = {'docx', 'xlsx', 'pptx'}
SIZE_LIMIT_TEXT   = 5  * 1024 * 1024   # 5 MB
SIZE_LIMIT_BINARY = 20 * 1024 * 1024   # 20 MB


def generate_title(path: Path) -> tuple[str, str]:
    """
    ファイルのタイトルを生成する。
    戻り値: (title, source)  source は 'content' または 'filename'
    """
    ext = path.suffix.lstrip('.').lower()
    base = path.stem
    size = path.stat().st_size

    title: str | None = None
    source = 'filename'

    if ext in TEXT_EXTS and size < SIZE_LIMIT_TEXT:
        try:
            with open(path, encoding='utf-8', errors='ignore') as f:
                content = f.read()
            extracted = extract_title_from_text(content, ext)
            if extracted and extracted.strip():
                title = extracted.strip()
                source = 'content'
        except Exception:
            pass

    elif ext == 'pdf' and size < SIZE_LIMIT_BINARY:
        try:
            extracted = extract_title_from_pdf(path)
            if extracted and extracted.strip():
                title = extracted.strip()
                source = 'content'
        except Exception:
            pass

    elif ext in OFFICE_EXTS and size < SIZE_LIMIT_BINARY:
        try:
            extracted = extract_title_from_office(path, ext)
            if extracted and extracted.strip():
                title = extracted.strip()
                source = 'content'
        except Exception:
            pass

    if not title:
        title = clean_filename(base)
        source = 'filename'

    return title, source


# ---------------------------------------------------------------------------
# ユーティリティ
# ---------------------------------------------------------------------------

def sanitize_filename(name: str) -> str:
    """OS で使えない文字を除去してファイル名を安全にする"""
    result = re.sub(r'[\\/:*?"<>|]', '_', name)
    result = re.sub(r'\s+', ' ', result).strip().strip('.')
    return result[:200] or 'untitled'


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def collect_files(paths: list[str], recursive: bool) -> list[Path]:
    files: list[Path] = []
    for path_arg in paths:
        p = Path(path_arg)
        if p.is_file():
            files.append(p)
        elif p.is_dir():
            pattern = '**/*' if recursive else '*'
            files.extend(f for f in p.glob(pattern) if f.is_file())
        else:
            print(f'⚠  見つかりません: {path_arg}', file=sys.stderr)
    return files


def main() -> None:
    parser = argparse.ArgumentParser(
        description='ファイル内容からタイトルを抽出して自動リネームします',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    parser.add_argument('paths', nargs='+', help='リネームするファイルまたはディレクトリ')
    parser.add_argument('-y', '--yes',       action='store_true', help='確認なしで実行')
    parser.add_argument('-n', '--dry-run',   action='store_true', help='プレビューのみ（実際にはリネームしない）')
    parser.add_argument('-r', '--recursive', action='store_true', help='サブディレクトリを再帰的に処理')
    args = parser.parse_args()

    files = collect_files(args.paths, args.recursive)
    if not files:
        print('処理対象のファイルがありません', file=sys.stderr)
        sys.exit(1)

    print(f'\n{len(files)} 件のファイルを分析中...\n')

    # --- プレビュー生成 ---
    Plan = tuple[Path, str, str, str]  # (path, new_name, source, status)
    plans: list[Plan] = []

    for f in files:
        title, source = generate_title(f)
        new_name = sanitize_filename(title) + f.suffix
        new_path = f.parent / new_name

        if f.name == new_name:
            status = 'same'
        elif new_path.exists():
            status = 'exists'
        else:
            status = 'ok'

        plans.append((f, new_name, source, status))

    # --- プレビュー表示 ---
    ok_count = 0
    for f, new_name, source, status in plans:
        src_label = '[コンテンツ]' if source == 'content' else '[ファイル名]'
        if status == 'same':
            print(f'  ─  {f.name}')
            print(f'       変更なし')
        elif status == 'exists':
            print(f'  ✗  {f.name}')
            print(f'       → {new_name}  ⚠ 同名ファイルが既に存在します（スキップ）')
        else:
            ok_count += 1
            print(f'  ✓  {f.name}  {src_label}')
            print(f'       → {new_name}')

    if ok_count == 0:
        print('\nリネームが必要なファイルはありません')
        return

    print(f'\n{ok_count} 件をリネームします')

    if args.dry_run:
        print('（--dry-run モード: 実際にはリネームしません）')
        return

    if not args.yes:
        try:
            ans = input('\n続行しますか？ [y/N]: ').strip().lower()
        except (EOFError, KeyboardInterrupt):
            print('\nキャンセルしました')
            return
        if ans not in ('y', 'yes'):
            print('キャンセルしました')
            return

    # --- リネーム実行 ---
    print()
    done = 0
    for f, new_name, _, status in plans:
        if status != 'ok':
            continue
        try:
            f.rename(f.parent / new_name)
            print(f'  ✓  {f.name}')
            print(f'       → {new_name}')
            done += 1
        except Exception as e:
            print(f'  ✗  {f.name}: {e}', file=sys.stderr)

    print(f'\n✅ {done} 件のファイルをリネームしました')


if __name__ == '__main__':
    main()
