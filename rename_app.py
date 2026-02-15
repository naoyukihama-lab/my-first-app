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
"""

import re
import threading
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


# ---------------------------------------------------------------------------
# アプリ本体
# ---------------------------------------------------------------------------

class RenameApp(_BASE):

    def __init__(self):
        super().__init__()
        ctk.set_appearance_mode('system')
        ctk.set_default_color_theme('blue')

        self.title('ファイル自動リネーム')
        self.geometry('1060x620')
        self.minsize(780, 440)

        self._plans: list[dict] = []
        self._selected_paths: list[str] = []
        self._sort_var   = tk.StringVar(value='ファイル名順')
        self._theme_var  = tk.StringVar(value='システム')
        self._sidebar_open = False

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

        ctk.CTkButton(top, text='📂 フォルダ', command=self._select_folder,
                      width=110).pack(side='left', padx=(0, 4))
        ctk.CTkButton(top, text='📄 ファイル', command=self._select_files,
                      width=110).pack(side='left', padx=(0, 10))

        self._recursive_var = tk.BooleanVar(value=False)
        ctk.CTkCheckBox(top, text='サブフォルダ', variable=self._recursive_var,
                        command=self._on_recursive_toggle).pack(side='left', padx=(0, 10))

        hint = 'ここにドロップ / またはファイル・フォルダを選択してください' if _DND_AVAILABLE \
               else 'ファイルまたはフォルダを選択してください'
        self._path_label = ctk.CTkLabel(top, text=hint, text_color='gray', anchor='w')
        self._path_label.pack(side='left', fill='x', expand=True, padx=(0, 10))

        ctk.CTkButton(top, text='⚙  設定', command=self._toggle_sidebar,
                      width=80).pack(side='right')

        # --- 本体：テーブル + サイドバー ---
        body = ctk.CTkFrame(self, fg_color='transparent')
        body.pack(fill='both', expand=True, padx=12, pady=0)
        body.columnconfigure(0, weight=1)
        body.rowconfigure(0, weight=1)
        self._body = body

        # テーブル
        table_frame = ctk.CTkFrame(body, fg_color='transparent')
        table_frame.grid(row=0, column=0, sticky='nsew')

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

        # DnD
        if _DND_AVAILABLE:
            for w in (self, self._tree):
                w.drop_target_register(DND_FILES)
                w.dnd_bind('<<Drop>>', self._on_drop)

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

        self._status_label = ctk.CTkLabel(bottom, text='Ready',
                                           text_color='gray', anchor='w')
        self._status_label.pack(side='left', padx=(16, 0))

    def _make_sidebar(self) -> ctk.CTkFrame:
        """設定サイドバー（grid col=1 に収まる）"""
        sb = ctk.CTkFrame(self._body, width=190, corner_radius=8)

        ctk.CTkLabel(sb, text='設定',
                     font=ctk.CTkFont(size=14, weight='bold')).pack(
                         padx=12, pady=(14, 6))

        # 並び順
        ctk.CTkLabel(sb, text='並び順', anchor='w').pack(fill='x', padx=12, pady=(8, 2))
        ctk.CTkOptionMenu(
            sb, values=list(SORT_OPTIONS.keys()),
            variable=self._sort_var, command=self._on_sort_change,
            width=166,
        ).pack(padx=12, pady=(0, 10))

        # テーマ
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
        """ルートウィンドウの背景を CTk のテーマに合わせる"""
        mode = ctk.get_appearance_mode()
        bg = '#212121' if mode == 'Dark' else '#ebebeb'
        try:
            self.configure(bg=bg)
        except Exception:
            pass

    # ------------------------------------------------------------------
    # ファイル選択
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
        """並び順変更時にデータを再ソートして再描画"""
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
        self._set_status(f'{total} 件分析完了  /  {ok_count} 件リネーム可能')
        self._rename_btn.configure(state='normal' if ok_count > 0 else 'disabled')

    def _set_status(self, text: str, color: str = 'gray'):
        self._status_label.configure(text=text, text_color=color)


# ---------------------------------------------------------------------------
# エントリポイント
# ---------------------------------------------------------------------------

if __name__ == '__main__':
    app = RenameApp()
    app.mainloop()
