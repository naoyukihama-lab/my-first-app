#!/usr/bin/env python3
"""
rename_app.py - GUI版ファイル自動リネームアプリ（CustomTkinter）

使い方:
  python rename_app.py

必要なもの:
  - Python 3.10 以上
  - pip install customtkinter
  - rename_files.py と同じフォルダに配置

オプションライブラリ:
  pip install pypdf        # PDF テキスト抽出強化
  pip install olefile      # 旧 Office (.doc/.xls/.ppt) 抽出強化
  pip install tkinterdnd2  # ドラッグ＆ドロップ対応
  pip install watchdog     # ホットフォルダ監視強化（なくてもポーリングで動作）
"""

import re
import threading
import time
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import customtkinter as ctk

from rename_files import collect_files, generate_title, sanitize_filename

# ドラッグ＆ドロップ（オプション）
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD  # type: ignore
    _DND_AVAILABLE = True
except ImportError:
    _DND_AVAILABLE = False

# ホットフォルダ監視（オプション）
try:
    from watchdog.observers import Observer                      # type: ignore
    from watchdog.events import FileSystemEventHandler          # type: ignore
    _WATCHDOG_AVAILABLE = True
except ImportError:
    _WATCHDOG_AVAILABLE = False

_BASE = TkinterDnD.Tk if _DND_AVAILABLE else ctk.CTk

# ---------------------------------------------------------------------------
# 並び順の定義
# ---------------------------------------------------------------------------

SORT_OPTIONS: dict[str, object] = {
    'ファイル名順':         lambda f: f.name.lower(),
    '更新日時（新→古）':   lambda f: -f.stat().st_mtime,
    '更新日時（古→新）':   lambda f:  f.stat().st_mtime,
    '拡張子順':             lambda f: (f.suffix.lower(), f.name.lower()),
}

# ---------------------------------------------------------------------------
# 定数
# ---------------------------------------------------------------------------

COL_CHECK  = 'check'
COL_OLD    = 'old'
COL_NEW    = 'new'
COL_SOURCE = 'source'

ST_OK     = 'ok'
ST_SAME   = 'same'
ST_EXISTS = 'exists'

CHECK_ON  = '✓'
CHECK_OFF = '─'

MODES = ['フォルダ参照', 'ファイル選択', 'クラウド', '監視フォルダ']

HELP_TEXT = """\
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
　ファイル自動リネームアプリ  使い方ガイド
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

【基本的な使い方】

  1. 入力方式をモードボタンで選択してください
  2. アクションボタンを押してファイルを読み込みます
  3. 候補一覧を確認し、対象ファイルにチェックを入れます
  4. 「リネーム実行」ボタンを押して完了です

  ※ ファイルやフォルダをアプリにドロップしても読み込めます
     （tkinterdnd2 インストール時のみ）

──────────────────────────────────────────

【入力モードの説明】

  📂 フォルダ参照
      フォルダを選択して、その中にあるすべてのファイルを
      一括で処理します。「サブフォルダ」にチェックを入れると
      入れ子のフォルダも対象になります。

  📄 ファイル選択
      複数のファイルを個別に選んで処理します。

  ☁  クラウド
      Google Drive・Dropbox・OneDrive などのローカル同期
      フォルダを自動検出して選択できます。

  👁  監視フォルダ
      指定したフォルダを常時監視します。新しいファイルが
      追加されると自動的に候補一覧を更新します。

──────────────────────────────────────────

【ファイル名の生成ルール】

  ファイルの内容（メタデータ・テキスト）から
  タイトル・作成者・日付を取得して命名します。

  例：
    ・文書タイトル（作成者）2024年1月
    ・プロジェクト提案書（山田太郎→鈴木部長）2024-03-15
    ・会議議事録_2024-01-20

  取得できない場合は元のファイル名を元に命名します。

──────────────────────────────────────────

【対応ファイル形式】

  ドキュメント:
    PDF, Word (.docx/.doc), Excel (.xlsx/.xls),
    PowerPoint (.pptx/.ppt), EPUB, ODT/ODS/ODP,
    Pages/Numbers/Keynote

  テキスト:
    TXT, Markdown, HTML, XML, JSON, YAML,
    CSV, RTF, LaTeX, ICS, VCF, INI/CFG

  プログラムコード:
    Python, JavaScript, TypeScript, Java, C/C++,
    Go, Rust, Ruby, PHP, Swift, Kotlin, その他多数

  メール:
    EML (.eml), Outlook MSG (.msg)

  画像（EXIFメタデータ):
    JPEG, PNG, TIFF, HEIC, WebP など

  音声（ID3タグ）:
    MP3, FLAC, M4A, AAC, OGG, OPUS など

──────────────────────────────────────────

【オプションライブラリ（インストールで機能強化）】

  pip install pypdf       → PDF テキスト抽出精度向上
  pip install olefile     → 旧 Office 形式 (.doc/.xls) 対応
  pip install pillow      → 画像 EXIF 取得精度向上
  pip install mutagen     → 音声タグ取得精度向上
  pip install watchdog    → 監視フォルダのリアルタイム検知
  pip install tkinterdnd2 → ドラッグ＆ドロップ対応

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
"""


# ---------------------------------------------------------------------------
# クラウドストレージのローカル同期フォルダを検出
# ---------------------------------------------------------------------------

def _detect_cloud_folders() -> dict[str, Path]:
    """よく使われるクラウドストレージのローカル同期フォルダを検索"""
    home = Path.home()
    candidates: list[tuple[str, list[Path]]] = [
        ('Google Drive',  [home / 'Google Drive',
                           home / 'Google Drive (マイドライブ)',
                           home / 'GoogleDrive',
                           home / 'My Drive']),
        ('Dropbox',       [home / 'Dropbox']),
        ('OneDrive',      [home / 'OneDrive',
                           *list(home.glob('OneDrive - *'))]),
        ('Box',           [home / 'Box', home / 'Box Sync']),
        ('iCloud Drive',  [home / 'iCloud Drive',
                           home / 'Library' / 'Mobile Documents' / 'com~apple~CloudDocs']),
        ('pCloud',        [home / 'pCloudDrive', home / 'pCloud Drive']),
        ('Mega',          [home / 'MEGA', home / 'MEGAsync']),
        ('Sync.com',      [home / 'Sync']),
        ('Tresorit',      [home / 'Tresorit']),
    ]
    found: dict[str, Path] = {}
    for name, paths in candidates:
        for p in paths:
            try:
                if p.exists():
                    found[name] = p
                    break
            except Exception:
                pass
    return found


# ---------------------------------------------------------------------------
# watchdog イベントハンドラ
# ---------------------------------------------------------------------------

if _WATCHDOG_AVAILABLE:
    class _WatchHandler(FileSystemEventHandler):  # type: ignore[misc]
        def __init__(self, callback):
            super().__init__()
            self._cb = callback

        def on_created(self, event):
            if not event.is_directory:
                self._cb()

        def on_moved(self, event):
            if not event.is_directory:
                self._cb()


# ---------------------------------------------------------------------------
# アプリ本体
# ---------------------------------------------------------------------------

class RenameApp(_BASE):

    def __init__(self):
        super().__init__()
        ctk.set_appearance_mode('system')
        ctk.set_default_color_theme('blue')

        self.title('ファイル自動リネーム')
        self.geometry('1100x660')
        self.minsize(800, 500)

        self._plans: list[dict] = []
        self._selected_paths: list[str] = []
        self._sort_var   = tk.StringVar(value='ファイル名順')
        self._theme_var  = tk.StringVar(value='システム')
        self._mode_var   = tk.StringVar(value='フォルダ参照')
        self._sidebar_open = False

        # ホットフォルダ監視状態
        self._watching       = False
        self._watch_path:   Path | None = None
        self._watch_observer = None
        self._watch_seen:   set[Path] = set()

        self._build_ui()
        self._sync_root_bg()

    # ------------------------------------------------------------------
    # UI 構築
    # ------------------------------------------------------------------

    def _build_ui(self):
        self._setup_treeview_style()

        # --- トップバー ---
        top = ctk.CTkFrame(self, fg_color='transparent')
        top.pack(fill='x', padx=12, pady=(10, 4))

        # モード選択セグメントボタン
        self._mode_seg = ctk.CTkSegmentedButton(
            top,
            values=MODES,
            variable=self._mode_var,
            command=self._on_mode_change,
            width=420,
        )
        self._mode_seg.pack(side='left', padx=(0, 10))

        # アクションボタン（モードに応じて変化）
        self._action_btn = ctk.CTkButton(
            top, text='📂 フォルダを選択',
            command=self._do_action, width=160,
        )
        self._action_btn.pack(side='left', padx=(0, 10))

        # サブフォルダチェック
        self._recursive_var = tk.BooleanVar(value=False)
        self._recursive_chk = ctk.CTkCheckBox(
            top, text='サブフォルダ', variable=self._recursive_var,
            command=self._on_recursive_toggle,
        )
        self._recursive_chk.pack(side='left', padx=(0, 10))

        # 設定ボタン
        ctk.CTkButton(top, text='⚙  設定', command=self._toggle_sidebar,
                      width=80).pack(side='right')

        # ステータスラベル（パス表示）
        self._path_label = ctk.CTkLabel(
            top,
            text='ファイルをドロップ、またはモードを選択して読み込んでください' if _DND_AVAILABLE
                 else 'モードを選択してファイルを読み込んでください',
            text_color='gray', anchor='w',
        )
        self._path_label.pack(side='left', fill='x', expand=True, padx=(0, 10))

        # --- 本体：コンテンツエリア + サイドバー ---
        body = ctk.CTkFrame(self, fg_color='transparent')
        body.pack(fill='both', expand=True, padx=12, pady=0)
        body.columnconfigure(0, weight=1)
        body.rowconfigure(0, weight=1)
        self._body = body

        # ウェルカム／ヘルプパネル（起動時表示）
        self._welcome_frame = self._make_welcome_panel(body)
        self._welcome_frame.grid(row=0, column=0, sticky='nsew')

        # テーブルフレーム（ファイル読み込み後に表示）
        table_frame = ctk.CTkFrame(body, fg_color='transparent')
        self._table_frame = table_frame

        cols = (COL_CHECK, COL_OLD, COL_NEW, COL_SOURCE)
        self._tree = ttk.Treeview(table_frame, columns=cols, show='headings',
                                   selectmode='none', style='App.Treeview')

        self._tree.heading(COL_CHECK,  text='対象')
        self._tree.heading(COL_OLD,    text='現在のファイル名')
        self._tree.heading(COL_NEW,    text='新しいファイル名')
        self._tree.heading(COL_SOURCE, text='取得元')

        self._tree.column(COL_CHECK,  width=54,  minwidth=54,  stretch=False, anchor='center')
        self._tree.column(COL_OLD,    width=320, minwidth=150, stretch=True)
        self._tree.column(COL_NEW,    width=320, minwidth=150, stretch=True)
        self._tree.column(COL_SOURCE, width=100, minwidth=80,  stretch=False, anchor='center')

        self._tree.tag_configure(ST_OK,     foreground='#2ecc71')
        self._tree.tag_configure(ST_SAME,   foreground='#888888')
        self._tree.tag_configure(ST_EXISTS, foreground='#e74c3c')

        vsb = ttk.Scrollbar(table_frame, orient='vertical',   command=self._tree.yview)
        hsb = ttk.Scrollbar(table_frame, orient='horizontal', command=self._tree.xview)
        self._tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        hsb.pack(side='bottom', fill='x')
        vsb.pack(side='right',  fill='y')
        self._tree.pack(fill='both', expand=True)
        self._tree.bind('<Button-1>', self._on_cell_click)

        # 設定サイドバー（初期は非表示）
        self._sidebar = self._make_sidebar()

        # DnD（常時有効 ─ ウェルカムパネル・テーブル両方に登録）
        if _DND_AVAILABLE:
            for w in (self, self._tree):
                w.drop_target_register(DND_FILES)
                w.dnd_bind('<<Drop>>', self._on_drop)
            # ウェルカムフレーム内ウィジェットにも登録
            try:
                self._welcome_frame.drop_target_register(DND_FILES)
                self._welcome_frame.dnd_bind('<<Drop>>', self._on_drop)
            except Exception:
                pass

        # --- ボトムバー ---
        bottom = ctk.CTkFrame(self, fg_color='transparent')
        bottom.pack(fill='x', padx=12, pady=(4, 10))

        ctk.CTkButton(bottom, text='すべて選択', command=self._select_all,
                      width=100).pack(side='left', padx=(0, 4))
        ctk.CTkButton(bottom, text='すべて解除', command=self._deselect_all,
                      width=100).pack(side='left')

        self._rename_btn = ctk.CTkButton(
            bottom, text='リネーム実行  ▶', command=self._do_rename,
            state='disabled', width=160,
            fg_color='#27ae60', hover_color='#1e8449', text_color='white',
        )
        self._rename_btn.pack(side='right')

        self._status_label = ctk.CTkLabel(bottom, text='準備完了',
                                           text_color='gray', anchor='w')
        self._status_label.pack(side='left', padx=(16, 0))

    def _make_welcome_panel(self, parent) -> ctk.CTkFrame:
        """起動時に表示するウェルカム／説明パネル"""
        frame = ctk.CTkFrame(parent, fg_color='transparent')
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(1, weight=1)

        # DnDゾーン（目立つ枠）
        if _DND_AVAILABLE:
            dnd_zone = ctk.CTkFrame(
                frame, corner_radius=12,
                border_width=2, border_color='#3a7ebf',
                fg_color=('gray90', 'gray20'),
                height=90,
            )
            dnd_zone.grid(row=0, column=0, sticky='ew', padx=20, pady=(14, 8))
            dnd_zone.grid_propagate(False)
            dnd_label = ctk.CTkLabel(
                dnd_zone,
                text='📂  ここにファイルやフォルダをドロップしてください',
                font=ctk.CTkFont(size=15, weight='bold'),
                text_color=('gray30', 'gray80'),
            )
            dnd_label.place(relx=0.5, rely=0.5, anchor='center')
            try:
                dnd_zone.drop_target_register(DND_FILES)
                dnd_zone.dnd_bind('<<Drop>>', self._on_drop)
                dnd_label.drop_target_register(DND_FILES)
                dnd_label.dnd_bind('<<Drop>>', self._on_drop)
            except Exception:
                pass
        else:
            sep_label = ctk.CTkLabel(
                frame,
                text='モードを選択してファイルを読み込んでください',
                font=ctk.CTkFont(size=14),
                text_color='gray',
            )
            sep_label.grid(row=0, column=0, pady=(14, 8))

        # ヘルプテキスト
        textbox = ctk.CTkTextbox(
            frame, wrap='word',
            font=ctk.CTkFont(family='monospace', size=12),
            activate_scrollbars=True,
        )
        textbox.grid(row=1, column=0, sticky='nsew', padx=20, pady=(0, 8))
        textbox.insert('end', HELP_TEXT)
        textbox.configure(state='disabled')

        return frame

    def _make_sidebar(self) -> ctk.CTkFrame:
        """設定サイドバー"""
        sb = ctk.CTkFrame(self._body, width=190, corner_radius=8)

        ctk.CTkLabel(sb, text='設定',
                     font=ctk.CTkFont(size=14, weight='bold')).pack(
                         padx=12, pady=(14, 6))

        ctk.CTkLabel(sb, text='並び順', anchor='w').pack(fill='x', padx=12, pady=(8, 2))
        ctk.CTkOptionMenu(
            sb, values=list(SORT_OPTIONS.keys()),
            variable=self._sort_var, command=self._on_sort_change,
            width=166,
        ).pack(padx=12, pady=(0, 10))

        ctk.CTkLabel(sb, text='テーマ', anchor='w').pack(fill='x', padx=12, pady=(4, 2))
        ctk.CTkSegmentedButton(
            sb, values=['ライト', 'ダーク', 'システム'],
            variable=self._theme_var, command=self._on_theme_change,
            width=166,
        ).pack(padx=12, pady=(0, 16))

        return sb

    def _setup_treeview_style(self):
        style = ttk.Style()
        style.theme_use('clam')
        style.configure('App.Treeview',         rowheight=28, font=('', 10))
        style.configure('App.Treeview.Heading', font=('', 9, 'bold'))

    # ------------------------------------------------------------------
    # モード切り替え
    # ------------------------------------------------------------------

    def _on_mode_change(self, mode: str):
        """モードに応じてアクションボタンとUI要素を更新"""
        # 監視フォルダ以外に切り替えたとき、監視中なら自動停止
        if self._watching and mode != '監視フォルダ':
            self._stop_watch()

        if mode == 'フォルダ参照':
            self._action_btn.configure(text='📂 フォルダを選択')
            self._recursive_chk.configure(state='normal')
        elif mode == 'ファイル選択':
            self._action_btn.configure(text='📄 ファイルを選択')
            self._recursive_chk.configure(state='disabled')
        elif mode == 'クラウド':
            self._action_btn.configure(text='☁  クラウドを選択')
            self._recursive_chk.configure(state='normal')
        elif mode == '監視フォルダ':
            if self._watching:
                self._action_btn.configure(
                    text='⏹ 監視を停止',
                    fg_color='#c0392b', hover_color='#922b21',
                )
            else:
                self._action_btn.configure(
                    text='👁 監視を開始',
                    fg_color=ctk.ThemeManager.theme['CTkButton']['fg_color'],
                    hover_color=ctk.ThemeManager.theme['CTkButton']['hover_color'],
                )
            self._recursive_chk.configure(state='disabled')

    def _do_action(self):
        """現在のモードに応じた処理を実行"""
        mode = self._mode_var.get()
        if mode == 'フォルダ参照':
            self._select_folder()
        elif mode == 'ファイル選択':
            self._select_files()
        elif mode == 'クラウド':
            self._select_cloud()
        elif mode == '監視フォルダ':
            self._toggle_watch()

    # ------------------------------------------------------------------
    # サイドバー開閉
    # ------------------------------------------------------------------

    def _toggle_sidebar(self):
        if self._sidebar_open:
            self._sidebar.grid_forget()
        else:
            self._sidebar.grid(row=0, column=1, sticky='ns', padx=(6, 0))
        self._sidebar_open = not self._sidebar_open

    # ------------------------------------------------------------------
    # 設定変更ハンドラ
    # ------------------------------------------------------------------

    def _on_sort_change(self, _=None):
        if self._plans:
            self._re_sort_and_repopulate()

    def _on_theme_change(self, value: str):
        mapping = {'ライト': 'light', 'ダーク': 'dark', 'システム': 'system'}
        ctk.set_appearance_mode(mapping.get(value, 'system'))
        self.after(50, self._sync_root_bg)

    def _sync_root_bg(self):
        mode = ctk.get_appearance_mode()
        bg = '#212121' if mode == 'Dark' else '#ebebeb'
        try:
            self.configure(bg=bg)
        except Exception:
            pass

    # ------------------------------------------------------------------
    # フォルダ参照 / ファイル参照
    # ------------------------------------------------------------------

    def _select_folder(self):
        d = filedialog.askdirectory(title='フォルダを選択')
        if not d:
            return
        self._selected_paths = [d]
        self._path_label.configure(text=d, text_color=self._text_color())
        self._analyze()

    def _select_files(self):
        files = filedialog.askopenfilenames(title='ファイルを選択')
        if not files:
            return
        self._selected_paths = list(files)
        self._path_label.configure(text=f'{len(files)} ファイル選択済み',
                                    text_color=self._text_color())
        self._analyze()

    # ------------------------------------------------------------------
    # クラウドストレージ選択
    # ------------------------------------------------------------------

    def _select_cloud(self):
        found = _detect_cloud_folders()
        if not found:
            messagebox.showinfo(
                'クラウドストレージ',
                'クラウドストレージの同期フォルダが検出されませんでした。\n'
                'Google Drive / Dropbox / OneDrive などのデスクトップアプリを\n'
                'インストールして同期しているか確認してください。\n\n'
                'フォルダを手動で選択します。',
            )
            self._select_folder()
            return
        self._show_cloud_dialog(found)

    def _show_cloud_dialog(self, found: dict[str, Path]):
        dlg = ctk.CTkToplevel(self)
        dlg.title('クラウドストレージを選択')
        dlg.geometry('400x320')
        dlg.resizable(False, False)
        dlg.grab_set()
        dlg.focus_set()

        ctk.CTkLabel(dlg, text='☁  クラウドストレージを選択',
                     font=ctk.CTkFont(size=14, weight='bold')).pack(
                         padx=20, pady=(16, 10))

        selected_var = tk.StringVar(value=str(next(iter(found.values()))))

        scroll_frame = ctk.CTkScrollableFrame(dlg, height=160)
        scroll_frame.pack(fill='x', padx=20, pady=(0, 10))

        for name, path in found.items():
            row = ctk.CTkFrame(scroll_frame, fg_color='transparent')
            row.pack(fill='x', pady=2)
            ctk.CTkRadioButton(
                row,
                text=f'{name}',
                variable=selected_var,
                value=str(path),
                font=ctk.CTkFont(size=12, weight='bold'),
            ).pack(anchor='w')
            ctk.CTkLabel(row, text=str(path), text_color='gray',
                         font=ctk.CTkFont(size=10)).pack(anchor='w', padx=(24, 0))

        btn_frame = ctk.CTkFrame(dlg, fg_color='transparent')
        btn_frame.pack(fill='x', padx=20, pady=(0, 16))

        def on_browse():
            dlg.destroy()
            self._select_folder()

        def on_ok():
            path_str = selected_var.get()
            if path_str:
                self._selected_paths = [path_str]
                label = f'☁  {Path(path_str).name}  ({path_str})'
                self._path_label.configure(text=label, text_color=self._text_color())
                dlg.destroy()
                self._analyze()

        ctk.CTkButton(btn_frame, text='📂 その他を参照...', command=on_browse,
                      width=140).pack(side='left')
        ctk.CTkButton(btn_frame, text='選択', command=on_ok,
                      width=100, fg_color='#27ae60',
                      hover_color='#1e8449').pack(side='right')

    # ------------------------------------------------------------------
    # ホットフォルダ（監視フォルダ）
    # ------------------------------------------------------------------

    def _toggle_watch(self):
        if self._watching:
            self._stop_watch()
        else:
            self._start_watch()

    def _start_watch(self):
        d = filedialog.askdirectory(title='監視するフォルダを選択')
        if not d:
            return
        self._watch_path = Path(d)
        self._watching   = True

        self._selected_paths = [d]
        self._path_label.configure(
            text=f'👁  監視中: {d}', text_color='orange')
        self._action_btn.configure(
            text='⏹ 監視を停止',
            fg_color='#c0392b', hover_color='#922b21')
        self._analyze()

        if _WATCHDOG_AVAILABLE:
            self._start_watchdog()
        else:
            self._start_polling()

    def _start_watchdog(self):
        handler = _WatchHandler(lambda: self.after(0, self._on_watch_changed))
        obs = Observer()
        obs.schedule(handler, str(self._watch_path), recursive=False)
        obs.start()
        self._watch_observer = obs

    def _start_polling(self):
        """watchdog がない場合は 2 秒おきにポーリング"""
        try:
            self._watch_seen = {
                f for f in self._watch_path.iterdir() if f.is_file()  # type: ignore[union-attr]
            }
        except Exception:
            self._watch_seen = set()

        def poll():
            while self._watching:
                time.sleep(2)
                try:
                    current = {
                        f for f in self._watch_path.iterdir() if f.is_file()  # type: ignore[union-attr]
                    }
                    if current != self._watch_seen:
                        self._watch_seen = current
                        self.after(0, self._on_watch_changed)
                except Exception:
                    pass

        threading.Thread(target=poll, daemon=True).start()

    def _on_watch_changed(self):
        """監視フォルダに変化があったとき → 自動再分析"""
        if not self._watching or not self._watch_path:
            return
        self._selected_paths = [str(self._watch_path)]
        self._analyze()

    def _stop_watch(self):
        self._watching = False
        if self._watch_observer is not None:
            try:
                self._watch_observer.stop()
                self._watch_observer.join(timeout=2)
            except Exception:
                pass
            self._watch_observer = None
        # アクションボタンの更新は「監視フォルダ」モードのときだけ行う
        # （他モードへ切り替えた場合は _on_mode_change 側でボタンを更新する）
        if self._mode_var.get() == '監視フォルダ':
            try:
                self._action_btn.configure(
                    text='👁 監視を開始',
                    fg_color=ctk.ThemeManager.theme['CTkButton']['fg_color'],
                    hover_color=ctk.ThemeManager.theme['CTkButton']['hover_color'],
                )
            except Exception:
                pass
        try:
            self._set_status('監視停止', 'gray')
            self._path_label.configure(text='監視停止', text_color='gray')
        except Exception:
            pass

    # ------------------------------------------------------------------
    # DnD
    # ------------------------------------------------------------------

    def _on_recursive_toggle(self):
        if self._selected_paths:
            self._analyze()

    def _on_drop(self, event) -> None:
        raw: str = event.data
        paths: list[str] = re.findall(r'\{([^}]+)\}', raw)
        remaining = re.sub(r'\{[^}]+\}', '', raw).strip()
        if remaining:
            paths.extend(remaining.split())
        paths = [p for p in paths if p]
        if not paths:
            return
        self._selected_paths = paths
        label = f'{len(paths)} 件をドロップ' if len(paths) > 1 else paths[0]
        self._path_label.configure(text=label, text_color=self._text_color())
        self._analyze()

    def _text_color(self) -> str:
        return 'white' if ctk.get_appearance_mode() == 'Dark' else 'black'

    # ------------------------------------------------------------------
    # コンテンツエリアの切り替え（ウェルカム ↔ テーブル）
    # ------------------------------------------------------------------

    def _show_table(self):
        """テーブルを表示し、ウェルカムパネルを隠す"""
        self._welcome_frame.grid_remove()
        self._table_frame.grid(row=0, column=0, sticky='nsew')

    def _show_welcome(self):
        """ウェルカムパネルを表示し、テーブルを隠す"""
        self._table_frame.grid_remove()
        self._welcome_frame.grid(row=0, column=0, sticky='nsew')

    # ------------------------------------------------------------------
    # 分析（バックグラウンドスレッド）
    # ------------------------------------------------------------------

    def _sorted_files(self, files: list[Path]) -> list[Path]:
        key = SORT_OPTIONS.get(self._sort_var.get(), lambda f: f.name.lower())
        try:
            return sorted(files, key=key)  # type: ignore[arg-type]
        except Exception:
            return files

    def _analyze(self):
        if not self._selected_paths:
            return
        self._set_status('分析中...', 'dodgerblue')
        self._rename_btn.configure(state='disabled')
        self._show_table()
        self._tree.delete(*self._tree.get_children())
        self._plans.clear()

        paths     = self._selected_paths[:]
        recursive = self._recursive_var.get()

        def worker():
            files = self._sorted_files(collect_files(paths, recursive))
            plans = []
            for f in files:
                title, source = generate_title(f)
                new_name = sanitize_filename(title) + f.suffix
                new_path = f.parent / new_name
                if f.name == new_name:
                    status = ST_SAME
                elif new_path.exists():
                    status = ST_EXISTS
                else:
                    status = ST_OK
                plans.append({
                    'path':     f,
                    'new_name': new_name,
                    'source':   source,
                    'status':   status,
                    'selected': status == ST_OK,
                })
            self.after(0, lambda: self._populate(plans))

        threading.Thread(target=worker, daemon=True).start()

    def _populate(self, plans: list[dict]):
        self._plans = plans
        if not plans:
            self._show_welcome()
            self._set_status('対象ファイルが見つかりませんでした', 'gray')
            return
        for p in plans:
            check = CHECK_ON if p['selected'] else CHECK_OFF
            if p['status'] == ST_SAME:
                src_label = '─'
            elif p['status'] == ST_EXISTS:
                src_label = '⚠ 同名あり'
            else:
                src_label = 'コンテンツ' if p['source'] == 'content' else 'ファイル名'
            self._tree.insert('', 'end',
                              values=(check, p['path'].name, p['new_name'], src_label),
                              tags=(p['status'],))
        ok_count = sum(1 for p in plans if p['selected'])
        self._update_status_and_btn(len(plans), ok_count)

    def _re_sort_and_repopulate(self):
        key = SORT_OPTIONS.get(self._sort_var.get(), lambda f: f.name.lower())
        try:
            self._plans.sort(key=lambda p: key(p['path']))  # type: ignore[operator]
        except Exception:
            pass
        self._tree.delete(*self._tree.get_children())
        self._populate(self._plans)

    # ------------------------------------------------------------------
    # チェックボックストグル
    # ------------------------------------------------------------------

    def _on_cell_click(self, event: tk.Event):
        if self._tree.identify_region(event.x, event.y) != 'cell':
            return
        if self._tree.identify_column(event.x) != '#1':
            return
        row_id = self._tree.identify_row(event.y)
        if not row_id:
            return
        plan = self._plans[self._tree.index(row_id)]
        if plan['status'] != ST_OK:
            return
        plan['selected'] = not plan['selected']
        vals = list(self._tree.item(row_id, 'values'))
        vals[0] = CHECK_ON if plan['selected'] else CHECK_OFF
        self._tree.item(row_id, values=vals)
        ok_count = sum(1 for p in self._plans if p['selected'])
        self._update_status_and_btn(len(self._plans), ok_count)

    def _select_all(self):
        self._set_check_all(True)

    def _deselect_all(self):
        self._set_check_all(False)

    def _set_check_all(self, value: bool):
        for row_id, plan in zip(self._tree.get_children(), self._plans):
            if plan['status'] != ST_OK:
                continue
            plan['selected'] = value
            vals = list(self._tree.item(row_id, 'values'))
            vals[0] = CHECK_ON if value else CHECK_OFF
            self._tree.item(row_id, values=vals)
        ok_count = sum(1 for p in self._plans if p['selected'])
        self._update_status_and_btn(len(self._plans), ok_count)

    # ------------------------------------------------------------------
    # リネーム実行
    # ------------------------------------------------------------------

    def _do_rename(self):
        targets = [p for p in self._plans if p['selected']]
        if not targets:
            return
        if not messagebox.askyesno('確認',
                                   f'{len(targets)} 件のファイルをリネームします。\n続行しますか？'):
            return
        done = 0
        errors: list[str] = []
        for p in targets:
            try:
                p['path'].rename(p['path'].parent / p['new_name'])
                done += 1
            except Exception as e:
                errors.append(f'{p["path"].name}: {e}')

        msg = f'{done} 件のリネームが完了しました。'
        if errors:
            msg += f'\n\nエラー ({len(errors)} 件):\n' + '\n'.join(errors[:10])
        messagebox.showinfo('完了', msg)
        self._analyze()

    # ------------------------------------------------------------------
    # ステータス
    # ------------------------------------------------------------------

    def _update_status_and_btn(self, total: int, ok_count: int):
        watch_mark = '  👁 監視中' if self._watching else ''
        self._set_status(f'{total} 件分析完了  /  {ok_count} 件リネーム可能{watch_mark}')
        self._rename_btn.configure(state='normal' if ok_count > 0 else 'disabled')

    def _set_status(self, text: str, color: str = 'gray'):
        self._status_label.configure(text=text, text_color=color)

    # ------------------------------------------------------------------
    # 終了時処理
    # ------------------------------------------------------------------

    def destroy(self):
        self._stop_watch()
        super().destroy()


# ---------------------------------------------------------------------------
# エントリポイント
# ---------------------------------------------------------------------------

if __name__ == '__main__':
    app = RenameApp()
    app.mainloop()
