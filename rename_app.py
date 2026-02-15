#!/usr/bin/env python3
"""
rename_app.py - GUI版ファイル自動リネームアプリ（Tkinter）

使い方:
  python rename_app.py

必要なもの:
  - Python 3.10 以上（tkinter は標準添付）
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

from rename_files import collect_files, generate_title, sanitize_filename

# ドラッグ＆ドロップ（オプション）
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD  # type: ignore
    _DND_AVAILABLE = True
except ImportError:
    _DND_AVAILABLE = False

_TkBase = TkinterDnD.Tk if _DND_AVAILABLE else tk.Tk

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

COLOR_OK     = '#1a7f37'
COLOR_SAME   = '#888888'
COLOR_EXISTS = '#cf222e'
COLOR_BG     = '#f6f8fa'


# ---------------------------------------------------------------------------
# アプリ本体
# ---------------------------------------------------------------------------

class RenameApp(_TkBase):

    def __init__(self):
        super().__init__()
        self.title('ファイル自動リネーム')
        self.geometry('960x560')
        self.minsize(720, 400)
        self.configure(bg=COLOR_BG)

        self._plans: list[dict] = []
        self._selected_paths: list[str] = []

        self._build_ui()

    # ------------------------------------------------------------------
    # UI 構築
    # ------------------------------------------------------------------

    def _build_ui(self):
        style = ttk.Style(self)
        style.theme_use('clam')
        style.configure('TFrame',     background=COLOR_BG)
        style.configure('TLabel',     background=COLOR_BG)
        style.configure('TCheckbutton', background=COLOR_BG)
        style.configure('Accent.TButton', font=('', 10, 'bold'))
        style.configure('Treeview',       rowheight=26)
        style.configure('Treeview.Heading', font=('', 9, 'bold'))

        # --- トップ：パス選択 ---
        top = ttk.Frame(self, padding=(10, 8, 10, 4))
        top.pack(fill='x', side='top')

        ttk.Button(top, text='📂  フォルダを選択', command=self._select_folder,
                   width=18).pack(side='left', padx=(0, 4))
        ttk.Button(top, text='📄  ファイルを選択', command=self._select_files,
                   width=18).pack(side='left', padx=(0, 8))

        self._recursive_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(top, text='サブフォルダも対象', variable=self._recursive_var,
                        command=self._on_recursive_toggle).pack(side='left', padx=(0, 8))

        hint = 'ここにドロップ可 / フォルダまたはファイルを選択してください' if _DND_AVAILABLE \
               else 'フォルダまたはファイルを選択してください'
        self._path_label = ttk.Label(top, text=hint, foreground='gray')
        self._path_label.pack(side='left', fill='x', expand=True)

        # --- 中段：テーブル ---
        table_frame = ttk.Frame(self, padding=(10, 0, 10, 0))
        table_frame.pack(fill='both', expand=True)

        cols = (COL_CHECK, COL_OLD, COL_NEW, COL_SOURCE)
        self._tree = ttk.Treeview(table_frame, columns=cols, show='headings',
                                   selectmode='none')

        self._tree.heading(COL_CHECK,  text='対象')
        self._tree.heading(COL_OLD,    text='現在のファイル名')
        self._tree.heading(COL_NEW,    text='新しいファイル名')
        self._tree.heading(COL_SOURCE, text='取得元')

        self._tree.column(COL_CHECK,  width=54,  minwidth=54,  stretch=False, anchor='center')
        self._tree.column(COL_OLD,    width=310, minwidth=150, stretch=True)
        self._tree.column(COL_NEW,    width=310, minwidth=150, stretch=True)
        self._tree.column(COL_SOURCE, width=100, minwidth=80,  stretch=False, anchor='center')

        self._tree.tag_configure(ST_OK,     foreground=COLOR_OK)
        self._tree.tag_configure(ST_SAME,   foreground=COLOR_SAME)
        self._tree.tag_configure(ST_EXISTS, foreground=COLOR_EXISTS)

        vsb = ttk.Scrollbar(table_frame, orient='vertical',   command=self._tree.yview)
        hsb = ttk.Scrollbar(table_frame, orient='horizontal', command=self._tree.xview)
        self._tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        hsb.pack(side='bottom', fill='x')
        vsb.pack(side='right',  fill='y')
        self._tree.pack(fill='both', expand=True)

        self._tree.bind('<Button-1>', self._on_cell_click)

        # ドラッグ＆ドロップ登録
        if _DND_AVAILABLE:
            for widget in (self, self._tree):
                widget.drop_target_register(DND_FILES)
                widget.dnd_bind('<<Drop>>', self._on_drop)

        # --- ボトム：操作ボタン + ステータス ---
        bottom = ttk.Frame(self, padding=(10, 6, 10, 8))
        bottom.pack(fill='x', side='bottom')

        ttk.Button(bottom, text='すべて選択', command=self._select_all,
                   width=12).pack(side='left', padx=(0, 4))
        ttk.Button(bottom, text='すべて解除', command=self._deselect_all,
                   width=12).pack(side='left')

        self._rename_btn = ttk.Button(bottom, text='リネーム実行  ▶',
                                      command=self._do_rename,
                                      style='Accent.TButton',
                                      state='disabled', width=18)
        self._rename_btn.pack(side='right')

        self._status_var = tk.StringVar(value='Ready')
        ttk.Label(bottom, textvariable=self._status_var,
                  foreground='gray').pack(side='left', padx=(16, 0))

    # ------------------------------------------------------------------
    # ファイル選択
    # ------------------------------------------------------------------

    def _select_folder(self):
        d = filedialog.askdirectory(title='フォルダを選択')
        if not d:
            return
        self._selected_paths = [d]
        self._path_label.config(text=d, foreground='black')
        self._analyze()

    def _select_files(self):
        files = filedialog.askopenfilenames(title='ファイルを選択')
        if not files:
            return
        self._selected_paths = list(files)
        self._path_label.config(text=f'{len(files)} ファイル選択済み', foreground='black')
        self._analyze()

    def _on_recursive_toggle(self):
        if self._selected_paths:
            self._analyze()

    def _on_drop(self, event) -> None:
        """ドラッグ＆ドロップされたファイル／フォルダを受け取る"""
        raw: str = event.data
        # tkinterdnd2 はスペースを含むパスを {} で囲んで返す
        paths: list[str] = re.findall(r'\{([^}]+)\}', raw)
        remaining = re.sub(r'\{[^}]+\}', '', raw).strip()
        if remaining:
            paths.extend(remaining.split())
        paths = [p for p in paths if p]
        if not paths:
            return
        self._selected_paths = paths
        label = f'{len(paths)} 件をドロップ' if len(paths) > 1 else paths[0]
        self._path_label.config(text=label, foreground='black')
        self._analyze()

    # ------------------------------------------------------------------
    # 分析（バックグラウンドスレッド）
    # ------------------------------------------------------------------

    def _analyze(self):
        if not self._selected_paths:
            return
        self._set_status('分析中...', 'blue')
        self._rename_btn.config(state='disabled')
        self._tree.delete(*self._tree.get_children())
        self._plans.clear()

        paths = self._selected_paths[:]
        recursive = self._recursive_var.get()

        def worker():
            files = collect_files(paths, recursive)
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

    # ------------------------------------------------------------------
    # チェックボックストグル
    # ------------------------------------------------------------------

    def _on_cell_click(self, event: tk.Event):
        region = self._tree.identify_region(event.x, event.y)
        if region != 'cell':
            return
        col = self._tree.identify_column(event.x)
        if col != '#1':
            return
        row_id = self._tree.identify_row(event.y)
        if not row_id:
            return
        idx = self._tree.index(row_id)
        plan = self._plans[idx]
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

        # 同じフォルダを再分析して表示を更新
        self._analyze()

    # ------------------------------------------------------------------
    # ステータスバー更新
    # ------------------------------------------------------------------

    def _update_status_and_btn(self, total: int, ok_count: int):
        self._set_status(f'{total} 件分析完了  /  {ok_count} 件リネーム可能', 'black')
        self._rename_btn.config(state='normal' if ok_count > 0 else 'disabled')

    def _set_status(self, text: str, color: str = 'black'):
        self._status_var.set(text)
        for w in self.pack_slaves():
            if isinstance(w, ttk.Frame):
                for child in w.pack_slaves():
                    if isinstance(child, ttk.Label) and child.cget('textvariable') == str(self._status_var):
                        child.config(foreground=color)


# ---------------------------------------------------------------------------
# エントリポイント
# ---------------------------------------------------------------------------

if __name__ == '__main__':
    app = RenameApp()
    app.mainloop()
