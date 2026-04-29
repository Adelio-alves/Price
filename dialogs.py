# -*- coding: utf-8 -*-
"""
dialogs.py
"""

import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from auth_service import find_user_by_password
from constants import ALL_COLUMNS, APP_TITLE, COLUMN_HEADINGS
from helpers import safe_str
from ui_components import CompactScrollFrame

try:
    from send2trash import send2trash  # noqa: F401
    SEND2TRASH_OK = True
except Exception:
    SEND2TRASH_OK = False


class ConfigDialog(tk.Toplevel):
    def __init__(self, master):
        super().__init__(master)
        self.title("Configurações")
        self.geometry("900x620")
        self.minsize(760, 520)
        self.transient(master)
        self.grab_set()
        master.center_window(self, master)

        outer = ttk.Frame(self, padding=10)
        outer.pack(fill="both", expand=True)

        ttk.Label(outer, text="Configurações", font=("Segoe UI", 12, "bold")).pack(anchor="w", pady=(0, 8))

        content = CompactScrollFrame(outer)
        content.pack(fill="both", expand=True)

        actions_box = ttk.LabelFrame(content.inner, text="Ações", padding=8)
        actions_box.pack(fill="x", pady=(0, 8))

        top_actions = ttk.Frame(actions_box)
        top_actions.pack(fill="x")

        ttk.Button(top_actions, text="Abrir pasta", command=master.open_folder).grid(row=0, column=0, padx=4, pady=4, sticky="ew")
        ttk.Button(top_actions, text="Salvar loja atual", command=master.save_current_store_file).grid(row=0, column=1, padx=4, pady=4, sticky="ew")
        ttk.Button(top_actions, text="Salvar tudo", command=master.save_all_files).grid(row=0, column=2, padx=4, pady=4, sticky="ew")
        ttk.Button(top_actions, text="Exportar XLSX", command=master.export_consolidated_xlsx).grid(row=1, column=0, padx=4, pady=4, sticky="ew")
        ttk.Button(top_actions, text="Gerar PDF", command=master.finalize_report).grid(row=1, column=1, padx=4, pady=4, sticky="ew")

        for i in range(3):
            top_actions.grid_columnconfigure(i, weight=1)

        options_box = ttk.LabelFrame(content.inner, text="Opções gerais", padding=8)
        options_box.pack(fill="x", pady=(0, 8))

        row1 = ttk.Frame(options_box)
        row1.pack(fill="x", pady=(0, 6))

        ttk.Label(row1, text="Filtro:").pack(side="left", padx=(0, 6))
        self.filter_var = tk.StringVar(value=master.filter_var.get())
        ttk.Combobox(
            row1,
            textvariable=self.filter_var,
            width=22,
            state="readonly",
            values=["TODOS", "SOMENTE ALTERADOS", "MARGEM BAIXA/NEGATIVA"]
        ).pack(side="left", padx=(0, 12))

        ttk.Label(row1, text="Buscar:").pack(side="left", padx=(0, 6))
        self.search_var = tk.StringVar(value=master.search_var.get())
        ttk.Entry(row1, textvariable=self.search_var, width=28).pack(side="left", padx=(0, 12))

        row2 = ttk.Frame(options_box)
        row2.pack(fill="x", pady=(0, 2))

        self.show_hidden_var = tk.BooleanVar(value=master.show_hidden_stores)
        ttk.Checkbutton(
            row2,
            text="Ver lojas editadas / já vistas",
            variable=self.show_hidden_var
        ).pack(side="left", padx=(0, 16))

        self.reopen_last_folder_var = tk.BooleanVar(value=bool(master.settings.get("reopen_last_folder_on_start", True)))
        ttk.Checkbutton(
            row2,
            text="Reabrir última pasta ao iniciar",
            variable=self.reopen_last_folder_var
        ).pack(side="left")

        folder_box = ttk.LabelFrame(content.inner, text="Pasta padrão", padding=8)
        folder_box.pack(fill="x", pady=(0, 8))

        row_folder = ttk.Frame(folder_box)
        row_folder.pack(fill="x")
        self.last_folder_var = tk.StringVar(value=master.settings.get("last_folder", ""))
        ttk.Label(row_folder, text="Pasta:").pack(side="left")
        ttk.Entry(row_folder, textvariable=self.last_folder_var).pack(side="left", fill="x", expand=True, padx=8)
        ttk.Button(row_folder, text="Selecionar", command=self.choose_last_folder).pack(side="left")

        post_box = ttk.LabelFrame(content.inner, text="Após gerar o PDF", padding=8)
        post_box.pack(fill="x", pady=(0, 8))

        self.move_enabled_var = tk.BooleanVar(value=bool(master.settings.get("post_pdf_move_enabled", False)))
        self.move_target_var = tk.StringVar(value=master.settings.get("post_pdf_target_folder", ""))
        self.trash_enabled_var = tk.BooleanVar(value=bool(master.settings.get("post_pdf_trash_enabled", False)))

        row_post1 = ttk.Frame(post_box)
        row_post1.pack(fill="x", pady=(0, 6))
        ttk.Checkbutton(
            row_post1,
            text="Mover arquivos para outra pasta após gerar o PDF",
            variable=self.move_enabled_var
        ).pack(side="left")

        row_post2 = ttk.Frame(post_box)
        row_post2.pack(fill="x", pady=(0, 6))
        ttk.Label(row_post2, text="Destino:").pack(side="left")
        ttk.Entry(row_post2, textvariable=self.move_target_var).pack(side="left", fill="x", expand=True, padx=8)
        ttk.Button(row_post2, text="Selecionar", command=self.choose_move_target).pack(side="left")

        row_post3 = ttk.Frame(post_box)
        row_post3.pack(fill="x")
        ttk.Checkbutton(
            row_post3,
            text="Enviar arquivos para a lixeira após gerar o PDF",
            variable=self.trash_enabled_var
        ).pack(side="left")
        ttk.Label(
            row_post3,
            text=" (send2trash ok)" if SEND2TRASH_OK else " (instale: pip install send2trash)",
            foreground="#6B7280"
        ).pack(side="left", padx=(8, 0))

        columns_box = ttk.LabelFrame(content.inner, text="Campos visíveis", padding=8)
        columns_box.pack(fill="x", pady=(0, 8))

        columns_wrap = ttk.Frame(columns_box)
        columns_wrap.pack(fill="x")

        self.col_vars = {}
        vis = master.get_field_visibility()

        cols_per_line = 3
        items = list(ALL_COLUMNS)
        for i, col in enumerate(items):
            row = i // cols_per_line
            col_pos = i % cols_per_line
            var = tk.BooleanVar(value=bool(vis.get(col, True)))
            self.col_vars[col] = var
            text = COLUMN_HEADINGS.get(col, col)
            chk = ttk.Checkbutton(columns_wrap, text=text, variable=var)
            chk.grid(row=row, column=col_pos, sticky="w", padx=10, pady=4)

        for i in range(cols_per_line):
            columns_wrap.grid_columnconfigure(i, weight=1)

        bottom = ttk.Frame(outer)
        bottom.pack(fill="x", pady=(8, 0))

        ttk.Button(bottom, text="Mostrar tudo", command=self.show_all_columns).pack(side="left")
        ttk.Button(bottom, text="Ocultar tudo", command=self.hide_all_columns).pack(side="left", padx=(8, 0))
        ttk.Button(bottom, text="Aplicar", command=self.apply).pack(side="right", padx=(8, 0))
        ttk.Button(bottom, text="Fechar", command=self.destroy).pack(side="right")

    def choose_last_folder(self):
        folder = filedialog.askdirectory(
            title="Selecione a pasta padrão",
            initialdir=self.last_folder_var.get() or os.getcwd()
        )
        if folder:
            self.last_folder_var.set(folder)

    def choose_move_target(self):
        folder = filedialog.askdirectory(
            title="Selecione a pasta de destino",
            initialdir=self.move_target_var.get() or os.getcwd()
        )
        if folder:
            self.move_target_var.set(folder)

    def show_all_columns(self):
        for var in self.col_vars.values():
            var.set(True)

    def hide_all_columns(self):
        for var in self.col_vars.values():
            var.set(False)

    def apply(self):
        if self.move_enabled_var.get() and not self.move_target_var.get().strip():
            messagebox.showwarning(APP_TITLE, "Defina a pasta de destino para mover os arquivos.")
            return

        if self.move_enabled_var.get() and self.trash_enabled_var.get():
            messagebox.showwarning(APP_TITLE, "Marque apenas uma opção: mover para pasta OU enviar para lixeira.")
            return

        if self.trash_enabled_var.get() and not SEND2TRASH_OK:
            messagebox.showwarning(APP_TITLE, "Para usar lixeira, instale: pip install send2trash")
            return

        visible_map = {k: bool(v.get()) for k, v in self.col_vars.items()}
        if not any(visible_map.values()):
            messagebox.showwarning(APP_TITLE, "Deixe pelo menos um campo visível.")
            return

        master = self.master
        master.filter_var.set(self.filter_var.get())
        master.search_var.set(self.search_var.get())
        master.show_hidden_stores = bool(self.show_hidden_var.get())

        master.settings["filter_mode"] = master.filter_var.get()
        master.settings["search_text"] = master.search_var.get()
        master.settings["show_hidden_stores"] = master.show_hidden_stores
        master.settings["reopen_last_folder_on_start"] = bool(self.reopen_last_folder_var.get())
        master.settings["last_folder"] = self.last_folder_var.get().strip()
        master.settings["field_visibility"] = visible_map
        master.settings["post_pdf_move_enabled"] = bool(self.move_enabled_var.get())
        master.settings["post_pdf_target_folder"] = self.move_target_var.get().strip()
        master.settings["post_pdf_trash_enabled"] = bool(self.trash_enabled_var.get())

        master.save_settings()
        master.apply_column_visibility()
        master.rebuild_store_listbox()
        master.select_current_store_in_listbox()
        master.refresh_table()
        self.destroy()


class FinalizeDialog(tk.Toplevel):
    def __init__(self, master, users, resumo_callback):
        super().__init__(master)
        self.title("Fechar relatório")
        self.users = users
        self.result = None
        self.resumo_callback = resumo_callback

        self.geometry("560x300")
        self.resizable(False, False)
        self.transient(master)
        self.grab_set()
        master.center_window(self, master)

        outer = ttk.Frame(self, padding=18)
        outer.pack(fill="both", expand=True)

        ttk.Label(outer, text="Resumo final", font=("Segoe UI", 14, "bold")).pack(anchor="w", pady=(0, 10))

        resumo = self.resumo_callback()
        info = (
            f"Lojas revisadas: {resumo['lojas_total']}\n"
            f"Lojas com alteração: {resumo['lojas_com_alteracao']}\n"
            f"Itens alterados: {resumo['itens_alterados']}\n"
            f"Agrupamentos finais: {resumo['agrupamentos']}"
        )
        ttk.Label(outer, text=info, justify="left").pack(anchor="w", pady=(0, 12))

        ttk.Separator(outer).pack(fill="x", pady=8)

        ttk.Label(
            outer,
            text="Digite a senha para fechar e assinar o relatório:",
            font=("Segoe UI", 10, "bold")
        ).pack(anchor="w", pady=(4, 6))

        self.password_var = tk.StringVar()
        ent = ttk.Entry(outer, textvariable=self.password_var, show="*", width=30)
        ent.pack(anchor="w")
        ent.focus_set()

        self.status_lbl = ttk.Label(outer, text="", foreground="red")
        self.status_lbl.pack(anchor="w", pady=(8, 0))

        btns = ttk.Frame(outer)
        btns.pack(fill="x", side="bottom", pady=(20, 0))
        ttk.Button(btns, text="Cancelar", command=self.destroy).pack(side="right", padx=(8, 0))
        ttk.Button(btns, text="Confirmar e gerar relatório", command=self.confirm).pack(side="right")

        self.bind("<Return>", lambda e: self.confirm())

    def confirm(self):
        senha = self.password_var.get().strip()
        if not senha:
            self.status_lbl.config(text="Digite a senha.")
            return

        user = find_user_by_password(self.users, senha)
        if not user:
            self.status_lbl.config(text="Senha inválida ou inativa.")
            return

        self.result = user
        self.destroy()