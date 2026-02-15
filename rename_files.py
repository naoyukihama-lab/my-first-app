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
  pip install olefile        # 旧 Office (.doc/.xls/.ppt) 抽出を強化（なくても動作）
"""

import argparse
import datetime
import json
import os
import re
import struct
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

    # RTF: {\info{\title ...}} ブロックからタイトルを抽出
    if ext == 'rtf':
        m = re.search(r'\\title ([^\\}{]+)', content)
        if not m:
            m = re.search(r'\\subject ([^\\}{]+)', content)
        if m:
            return m.group(1).strip()[:80]

    # LaTeX: \title{...} / \author{...}
    if ext in ('tex', 'latex', 'sty', 'cls'):
        m = re.search(r'\\title\{([^}]+)\}', content)
        if m:
            # LaTeX コマンドを除去: \textbf{foo} → foo
            t = re.sub(r'\\[a-zA-Z]+\{([^}]*)\}', r'\1', m.group(1))
            t = re.sub(r'\\[a-zA-Z]+', '', t).strip()
            if t:
                return t[:80]

    # iCalendar (.ics): SUMMARY: または X-WR-CALNAME:
    if ext in ('ics', 'vcs'):
        for line in lines:
            if line.startswith('SUMMARY:') or line.startswith('X-WR-CALNAME:'):
                return line.split(':', 1)[1].strip()[:80]

    # vCard (.vcf): FN: (Full Name)
    if ext == 'vcf':
        for line in lines:
            if line.startswith('FN:'):
                return line[3:].strip()[:80]
            if line.startswith('N:'):
                parts = [p.strip() for p in line[2:].split(';') if p.strip()]
                name = ' '.join(parts[:2])
                if name:
                    return name[:80]

    # 設定ファイル: [Section] または title/name キー
    if ext in ('ini', 'cfg', 'conf', 'config', 'properties', 'env'):
        for line in lines[:30]:
            m = re.match(r'^\[([^\]]+)\]', line)
            if m:
                sec = m.group(1).strip()
                if sec.lower() not in ('general', 'default', 'main', 'settings', 'global', 'common'):
                    return sec[:80]
        for line in lines[:50]:
            m = re.match(r'^(?:title|name|app_?name|project_?name)\s*[=:]\s*["\']?(.+?)["\']?\s*$',
                         line, re.IGNORECASE)
            if m and m.group(1).strip():
                return m.group(1).strip()[:80]

    # コードファイル: 先頭コメント → クラス名 → 関数名
    code_exts = {
        'js', 'mjs', 'cjs', 'ts', 'mts', 'tsx', 'jsx', 'vue', 'svelte',
        'py', 'pyw', 'pyi', 'rb', 'php', 'java', 'kt', 'swift', 'go', 'rs',
        'cpp', 'cxx', 'cc', 'c', 'h', 'hpp', 'cs', 'vb', 'fs',
        'sh', 'bash', 'zsh', 'fish', 'ps1', 'sql', 'r', 'lua',
        'css', 'scss', 'sass', 'less',
        'hs', 'ex', 'exs', 'clj', 'elm', 'jl', 'nim', 'zig', 'cr', 'd',
        'ml', 'scala', 'dart', 'pl', 'pm', 'tcl', 'awk', 'coffee',
        'pas', 'f90', 'for', 'erl', 'groovy', 'gradle', 'tf', 'hcl',
        'proto', 'graphql', 'gql', 'makefile', 'mk', 'dockerfile',
        'rkt', 'lisp', 'scm',
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

def extract_title_from_pdf(path: Path) -> tuple[str | None, str | None, str | None]:
    """PDF から (title, author, created) を抽出"""
    title = author = created = None

    # pypdf が入っていれば優先使用
    try:
        import pypdf  # type: ignore
        reader = pypdf.PdfReader(str(path))
        meta = reader.metadata
        if meta:
            if getattr(meta, 'title', None) and meta.title.strip():
                title = meta.title.strip()[:80]
            if getattr(meta, 'author', None) and meta.author.strip():
                author = meta.author.strip()[:50]
            raw_date = (getattr(meta, 'creation_date_raw', None)
                        or meta.get('/CreationDate', ''))
            if raw_date:
                created = _parse_pdf_date(str(raw_date))
        if not title and reader.pages:
            text = reader.pages[0].extract_text() or ''
            for line in text.split('\n'):
                line = line.strip()
                if len(line) >= 4:
                    title = line[:80]
                    break
    except ImportError:
        pass
    except Exception:
        pass

    # フォールバック: バイナリから /Title /Author /CreationDate を探す
    if not title:
        try:
            with open(path, 'rb') as f:
                raw = f.read(65536)
            text = raw.decode('latin-1', errors='ignore')
            m = re.search(r'/Title\s*\(([^)]{2,80})\)', text)
            if m:
                title = m.group(1).strip()[:80] or None
            if not author:
                m = re.search(r'/Author\s*\(([^)]{2,50})\)', text)
                if m:
                    author = m.group(1).strip()[:50] or None
            if not created:
                m = re.search(r'/CreationDate\s*\(([^)]+)\)', text)
                if m:
                    created = _parse_pdf_date(m.group(1))
        except Exception:
            pass

    return title, author, created


# ---------------------------------------------------------------------------
# Office Open XML（DOCX / XLSX / PPTX）からタイトル抽出
# ---------------------------------------------------------------------------

def extract_title_from_office(path: Path, ext: str) -> tuple[str | None, str | None, str | None]:
    """Office Open XML（DOCX/XLSX/PPTX）から (title, author, created) を抽出"""
    title = author = created = None
    try:
        with zipfile.ZipFile(path, 'r') as z:
            names = z.namelist()

            # docProps/core.xml から title / creator / created を取得
            if 'docProps/core.xml' in names:
                xml = z.read('docProps/core.xml').decode('utf-8', errors='ignore')
                m = re.search(r'<dc:title>([^<]+)</dc:title>', xml, re.IGNORECASE)
                if m and m.group(1).strip():
                    title = m.group(1).strip()[:80]
                m = re.search(r'<dc:creator>([^<]+)</dc:creator>', xml, re.IGNORECASE)
                if m and m.group(1).strip():
                    author = m.group(1).strip()[:50]
                m = re.search(r'<dcterms:created[^>]*>([^<]+)</dcterms:created>', xml, re.IGNORECASE)
                if m and m.group(1).strip():
                    created = _parse_iso_date(m.group(1).strip())

            # タイトルが取れなかった場合はコンテンツから補完
            if not title:
                if ext == 'docx' and 'word/document.xml' in names:
                    xml = z.read('word/document.xml').decode('utf-8', errors='ignore')
                    for para in re.findall(r'<w:p[ >].*?</w:p>', xml, re.DOTALL):
                        parts = re.findall(r'<w:t[^>]*>([^<]+)</w:t>', para)
                        text = ''.join(parts).strip()
                        if len(text) >= 4:
                            title = text[:80]
                            break
                elif ext == 'xlsx':
                    wb = next((n for n in names if re.match(r'xl/workbook\.xml', n)), None)
                    if wb:
                        xml = z.read(wb).decode('utf-8', errors='ignore')
                        m = re.search(r'<sheet[^>]+name="([^"]+)"', xml, re.IGNORECASE)
                        if m and m.group(1).strip():
                            title = m.group(1).strip()
                elif ext == 'pptx':
                    slides = sorted(n for n in names if re.match(r'ppt/slides/slide\d+\.xml', n))
                    if slides:
                        xml = z.read(slides[0]).decode('utf-8', errors='ignore')
                        m = re.search(r'<a:t>([^<]{2,})</a:t>', xml)
                        if m and m.group(1).strip():
                            title = m.group(1).strip()[:80]
    except Exception:
        pass

    return title, author, created


# ---------------------------------------------------------------------------
# 旧 Office バイナリ形式（.doc / .xls / .ppt）からタイトル抽出
# ---------------------------------------------------------------------------

def _filetime_to_date(filetime: int) -> str | None:
    """Windows FILETIME（100ns 単位）を 'YYYY年MM月DD日' に変換"""
    EPOCH_DIFF = 116444736000000000  # 1601-01-01 〜 1970-01-01 の 100ns 数
    try:
        unix_sec = (filetime - EPOCH_DIFF) / 10_000_000
        if unix_sec < 0 or unix_sec > 32503680000:
            return None
        dt = datetime.datetime(1970, 1, 1) + datetime.timedelta(seconds=unix_sec)
        return f'{dt.year}年{dt.month:02d}月{dt.day:02d}日'
    except Exception:
        return None


def _parse_iso_date(s: str) -> str | None:
    """ISO 8601 文字列（2024-01-15T...）を 'YYYY年MM月DD日' に変換"""
    m = re.match(r'(\d{4})-(\d{2})-(\d{2})', s)
    return f'{m.group(1)}年{m.group(2)}月{m.group(3)}日' if m else None


def _parse_pdf_date(s: str) -> str | None:
    """PDF 日付文字列（D:20240115...）を 'YYYY年MM月DD日' に変換"""
    m = re.match(r'D:(\d{4})(\d{2})(\d{2})', s)
    return f'{m.group(1)}年{m.group(2)}月{m.group(3)}日' if m else None


def _parse_prop_set_info(data: bytes) -> tuple[str | None, str | None, str | None]:
    """Windows PROPSET（SummaryInformation）から (title, author, created) を取得"""
    title = author = created = None
    try:
        if len(data) < 48:
            return title, author, created
        c_sections = struct.unpack_from('<I', data, 24)[0]
        if c_sections < 1:
            return title, author, created
        section_offset = struct.unpack_from('<I', data, 44)[0]
        if section_offset + 8 > len(data):
            return title, author, created
        prop_count = struct.unpack_from('<I', data, section_offset + 4)[0]

        def _read_str(abs_off: int) -> str | None:
            if abs_off + 8 > len(data):
                return None
            vt = struct.unpack_from('<I', data, abs_off)[0]
            if vt == 0x1E:  # VT_LPSTR
                n = struct.unpack_from('<I', data, abs_off + 4)[0]
                raw = data[abs_off + 8: abs_off + 8 + n].rstrip(b'\x00')
                return (raw.decode('cp932', errors='ignore')
                        or raw.decode('latin-1', errors='ignore')).strip() or None
            if vt == 0x1F:  # VT_LPWSTR
                n = struct.unpack_from('<I', data, abs_off + 4)[0]
                raw = data[abs_off + 8: abs_off + 8 + n * 2]
                return raw.decode('utf-16-le', errors='ignore').rstrip('\x00').strip() or None
            return None

        for i in range(min(prop_count, 200)):
            entry_off = section_offset + 8 + i * 8
            if entry_off + 8 > len(data):
                break
            pid, prop_off = struct.unpack_from('<II', data, entry_off)
            abs_off = section_offset + prop_off
            if pid == 2:    # PIDSI_TITLE
                title = _read_str(abs_off)
            elif pid == 4:  # PIDSI_AUTHOR
                author = _read_str(abs_off)
            elif pid == 12: # PIDSI_CREATE_DTM
                if abs_off + 12 <= len(data):
                    vt = struct.unpack_from('<I', data, abs_off)[0]
                    if vt == 0x40:  # VT_FILETIME
                        ft = struct.unpack_from('<Q', data, abs_off + 4)[0]
                        created = _filetime_to_date(ft)
    except Exception:
        pass
    return title, author, created


def _extract_xls_sheet_names(data: bytes) -> list[str]:
    """BIFF8 データストリームから BOUNDSHEET レコード（0x0085）でシート名を収集"""
    names: list[str] = []
    i = 0
    while i < len(data) - 4:
        try:
            rec_type = struct.unpack_from('<H', data, i)[0]
            rec_len  = struct.unpack_from('<H', data, i + 2)[0]
        except struct.error:
            break
        if rec_type == 0x0085 and rec_len >= 6:  # BOUNDSHEET
            rec = data[i + 4: i + 4 + rec_len]
            if len(rec) >= 6:
                name_len = rec[4]
                flag     = rec[5]
                raw      = rec[6:]
                if flag & 0x01:  # Unicode
                    name = raw[:name_len * 2].decode('utf-16-le', errors='ignore')
                else:            # Latin / MBCS
                    name = raw[:name_len].decode('cp932', errors='ignore') \
                           or raw[:name_len].decode('latin-1', errors='ignore')
                if name.strip():
                    names.append(name.strip())
        i += 4 + max(0, rec_len)
    return names


def _extract_ppt_texts(data: bytes) -> list[str]:
    """PowerPoint Document ストリームから TextCharsAtom / TextBytesAtom を収集"""
    texts: list[str] = []
    i = 0
    while i < len(data) - 8:
        try:
            rec_type = struct.unpack_from('<H', data, i + 2)[0]
            rec_len  = struct.unpack_from('<I', data, i + 4)[0]
        except struct.error:
            break
        if rec_len > 10 * 1024 * 1024:  # 異常に大きいレコードはスキップ
            i += 1
            continue
        if rec_type == 0x0FA0 and rec_len > 0:  # TextCharsAtom (UTF-16LE)
            raw  = data[i + 8: i + 8 + rec_len]
            text = raw.decode('utf-16-le', errors='ignore').rstrip('\x00')
            if text.strip():
                texts.append(text.strip())
        elif rec_type == 0x0FA8 and rec_len > 0:  # TextBytesAtom (Latin)
            raw  = data[i + 8: i + 8 + rec_len]
            text = raw.decode('latin-1', errors='ignore').rstrip('\x00')
            if text.strip():
                texts.append(text.strip())
        i += 8 + max(0, rec_len)
    return texts


def _scan_utf16le_strings(data: bytes, min_len: int = 6) -> list[str]:
    """バイナリデータから UTF-16LE テキスト文字列をスキャン（.doc フォールバック用）"""
    results: list[str] = []
    i = 0
    n = len(data)
    while i < n - 1:
        if 0x20 <= data[i] <= 0x7E and data[i + 1] == 0x00:
            j = i
            while j + 1 < n and 0x20 <= data[j] <= 0x7E and data[j + 1] == 0x00:
                j += 2
            char_count = (j - i) // 2
            if char_count >= min_len:
                text = data[i:j].decode('utf-16-le', errors='ignore').strip()
                if text:
                    results.append(text)
            i = j + 2
        else:
            i += 1
    return results


def extract_title_from_old_office(path: Path, ext: str) -> tuple[str | None, str | None, str | None]:
    """旧 Office 形式（.doc / .xls / .ppt）から (title, author, created) を抽出"""
    title = author = created = None

    # --- 1. olefile 経由（最も正確） ---
    try:
        import olefile  # type: ignore
        with olefile.OleFileIO(str(path)) as ole:
            si = '\x05SummaryInformation'
            if ole.exists(si):
                title, author, created = _parse_prop_set_info(ole.openstream(si).read())

            if not title:
                if ext == 'xls':
                    for stream in ('Workbook', 'Book'):
                        if ole.exists(stream):
                            names = _extract_xls_sheet_names(ole.openstream(stream).read())
                            if names:
                                title = names[0][:80]
                                break
                elif ext == 'ppt':
                    if ole.exists('PowerPoint Document'):
                        texts = _extract_ppt_texts(ole.openstream('PowerPoint Document').read())
                        for t in texts:
                            if len(t) >= 4:
                                title = t[:80]
                                break
                elif ext == 'doc':
                    if ole.exists('WordDocument'):
                        raw = ole.openstream('WordDocument').read()
                        for s in _scan_utf16le_strings(raw, min_len=6):
                            if re.match(r'^[A-Za-z\u3040-\u9FFF\u4E00-\u9FFF]', s) and len(s) >= 4:
                                title = s[:80]
                                break
    except ImportError:
        pass
    except Exception:
        pass

    if title:
        return title, author, created

    # --- 2. バイナリスキャン（olefile なし・フォールバック） ---
    try:
        with open(path, 'rb') as f:
            data = f.read()

        magic = b'\x05SummaryInformation'
        idx = data.find(magic)
        if idx != -1:
            chunk = data[idx + len(magic): idx + len(magic) + 4096]
            title, author, created = _parse_prop_set_info(chunk)

        if not title:
            if ext == 'xls':
                names = _extract_xls_sheet_names(data)
                if names:
                    title = names[0][:80]
            elif ext == 'ppt':
                texts = _extract_ppt_texts(data)
                for t in texts:
                    if len(t) >= 4:
                        title = t[:80]
                        break
            elif ext == 'doc':
                for s in _scan_utf16le_strings(data, min_len=6):
                    if re.match(r'^[A-Za-z\u3040-\u9FFF\u4E00-\u9FFF]', s) and len(s) >= 4:
                        title = s[:80]
                        break
    except Exception:
        pass

    return title, author, created


# ---------------------------------------------------------------------------
# タイトル生成メイン
# ---------------------------------------------------------------------------

TEXT_EXTS = {
    # ドキュメント・マークアップ
    'txt', 'md', 'markdown', 'rst', 'adoc', 'asciidoc', 'textile', 'wiki',
    'tex', 'latex', 'sty', 'cls', 'bib', 'rtf',
    # Web
    'html', 'htm', 'xhtml', 'css', 'scss', 'sass', 'less', 'stylus', 'styl',
    'svg',
    # テンプレート
    'astro', 'svelte', 'vue', 'jsx', 'tsx',
    'njk', 'j2', 'jinja', 'jinja2', 'hbs', 'handlebars', 'mustache',
    'pug', 'jade', 'twig', 'blade', 'erb', 'ejs',
    # データ・設定
    'json', 'jsonc', 'jsonl', 'ndjson',
    'xml', 'csv', 'tsv', 'yaml', 'yml', 'toml',
    'ini', 'cfg', 'conf', 'config', 'properties', 'env',
    'tf', 'hcl',                    # Terraform / HCL
    'proto', 'graphql', 'gql',      # Protocol Buffers / GraphQL
    'gradle', 'groovy',             # Gradle / Groovy
    # カレンダー・連絡先
    'ics', 'vcs', 'vcf',
    # プログラミング言語
    'js', 'mjs', 'cjs', 'ts', 'mts',
    'py', 'pyw', 'pyi',
    'rb', 'rake',
    'php', 'php3', 'php4', 'php5', 'phtml',
    'java', 'kt', 'kts', 'groovy',
    'swift', 'go', 'rs',
    'cpp', 'cxx', 'cc', 'c', 'h', 'hpp', 'hxx',
    'cs', 'vb', 'fs', 'fsx',       # C# / VB.NET / F#
    'sh', 'bash', 'zsh', 'fish', 'ksh', 'csh',
    'ps1', 'psm1', 'psd1',         # PowerShell
    'sql', 'ddl', 'dml',
    'r', 'rmd',                    # R
    'lua', 'm',                    # Lua / MATLAB・Objective-C
    'mm',                          # Objective-C++
    # 関数型・その他
    'hs', 'lhs',                   # Haskell
    'ex', 'exs',                   # Elixir
    'clj', 'cljs', 'cljc',        # Clojure
    'elm',                         # Elm
    'jl',                          # Julia
    'nim',                         # Nim
    'zig',                         # Zig
    'cr',                          # Crystal
    'rkt',                         # Racket
    'd',                           # D言語
    'ml', 'mli',                   # OCaml
    'scala', 'sc',                 # Scala
    'dart',                        # Dart
    'pl', 'pm',                    # Perl
    'tcl',                         # Tcl
    'awk',                         # AWK
    'coffee',                      # CoffeeScript
    'pas', 'pp',                   # Pascal
    'f90', 'f95', 'f03', 'for', 'f',   # Fortran
    'v', 'vhd', 'vhdl', 'sv',     # Verilog / VHDL
    'lisp', 'scm', 'ss',          # Lisp / Scheme
    'erl', 'hrl',                  # Erlang
    'makefile', 'mk', 'cmake',     # ビルドスクリプト
    'dockerfile',
}
OFFICE_EXTS     = {'docx', 'xlsx', 'pptx'}
OLD_OFFICE_EXTS = {'doc', 'xls', 'ppt'}
EPUB_EXTS       = {'epub'}
ODF_EXTS        = {'odt', 'ods', 'odp', 'odg', 'ott', 'ots', 'otp'}
IWORK_EXTS      = {'pages', 'numbers', 'keynote'}
IMAGE_EXTS      = {'jpg', 'jpeg', 'png', 'tiff', 'tif', 'webp', 'heic', 'heif', 'bmp', 'gif'}
AUDIO_EXTS      = {'mp3', 'flac', 'm4a', 'aac', 'ogg', 'opus', 'wma', 'aiff', 'ape', 'wav'}
VIDEO_EXTS      = {'mp4', 'mkv', 'avi', 'mov', 'wmv', 'm4v', 'webm', 'flv', 'mpg', 'mpeg', 'ts'}
EML_EXTS        = {'eml', 'mbox'}
MSG_EXTS        = {'msg'}
SIZE_LIMIT_TEXT   = 50  * 1024 * 1024   # 50 MB
SIZE_LIMIT_BINARY = 200 * 1024 * 1024   # 200 MB


def build_display_name(title: str, author: str | None, created: str | None) -> str:
    """'文書名（作成者）作成日付' の形式でファイル名を組み立てる"""
    name = title
    if author:
        name += f'（{author}）'
    if created:
        name += created
    return name


def generate_title(path: Path) -> tuple[str, str]:
    """
    ファイルのタイトルを生成する。
    戻り値: (display_name, source)  source は 'content' または 'filename'
    """
    ext = path.suffix.lstrip('.').lower()
    base = path.stem
    size = path.stat().st_size

    raw_title: str | None = None
    author:    str | None = None
    created:   str | None = None
    source = 'filename'

    if ext in TEXT_EXTS and size < SIZE_LIMIT_TEXT:
        try:
            with open(path, encoding='utf-8', errors='ignore') as f:
                content = f.read()
            extracted = extract_title_from_text(content, ext)
            if extracted and extracted.strip():
                raw_title = extracted.strip()
                source = 'content'
        except Exception:
            pass

    elif ext == 'pdf' and size < SIZE_LIMIT_BINARY:
        try:
            raw_title, author, created = extract_title_from_pdf(path)
            if raw_title:
                raw_title = raw_title.strip()
                source = 'content'
        except Exception:
            pass

    elif ext in OFFICE_EXTS and size < SIZE_LIMIT_BINARY:
        try:
            raw_title, author, created = extract_title_from_office(path, ext)
            if raw_title:
                raw_title = raw_title.strip()
                source = 'content'
        except Exception:
            pass

    elif ext in OLD_OFFICE_EXTS and size < SIZE_LIMIT_BINARY:
        try:
            raw_title, author, created = extract_title_from_old_office(path, ext)
            if raw_title:
                raw_title = raw_title.strip()
                source = 'content'
        except Exception:
            pass

    if not raw_title:
        raw_title = clean_filename(base)
        source = 'filename'

    return build_display_name(raw_title, author, created), source


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
