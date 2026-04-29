# -*- coding: utf-8 -*-
"""
app.py
"""

import gc
import json
import os
import shutil
import subprocess
import sys
import traceback
from collections import OrderedDict
from datetime import datetime

import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from auth_service import load_authorization_file
from constants import (
    ALL_COLUMNS,
    APP_TITLE,
    CACHE_MAX_STORES,
    COLUMN_HEADINGS,
    CONFIG_FILE,
    DEFAULT_FIELD_VISIBILITY,
    DEFAULT_SETTINGS,
    METRIC_CARD_PADDING,
    METRICS_WRAP_PADY,
    METRIC_VALUE_PADY,
    STORE_LIST_HEIGHT,
    STORE_PANEL_WIDTH_CHARS,
)
from dialogs import ConfigDialog, FinalizeDialog, SEND2TRASH_OK
from excel_service import (
    list_excel_files,
    load_full_store_data,
    scan_store_metadata,
)
from helpers import (
    float_to_br,
    format_product_code,
    maybe_number_to_br,
    money_to_float,
    normalize_text,
    percent_to_br,
    safe_str,
    sanitize_decimal_text_for_entry,
)
from pdf_service import build_pdf

try:
    from send2trash import send2trash
except Exception:
    send2trash = None


REPORT_FILTER_OPTIONS = (
    "TODOS",
    "PRECISAM DE ATENÇÃO",
    "MARGEM ABAIXO DE",
    "MARGEM ACIMA DE",
    "SOMENTE ALTERADOS",
)


def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)


def get_app_base_dir():
    try:
        return sys._MEIPASS
    except Exception:
        return os.path.dirname(os.path.abspath(__file__))


class PriceEditorApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)

        if sys.platform.startswith("win"):
            try:
                self.iconbitmap(resource_path("app.ico"))
            except Exception:
                pass

        if sys.platform.startswith("win"):
            try:
                self.state("zoomed")
            except Exception:
                self.geometry("1600x860")
        else:
            self.geometry("1600x860")

        self.minsize(1220, 700)

        self.base_dir = get_app_base_dir()
        self.config_path = os.path.join(self.base_dir, CONFIG_FILE)
        self.settings = self.load_settings()

        self.users = load_authorization_file(self.base_dir)

        self.stores_meta = []
        self.loaded_cache = OrderedDict()
        self.edits_map = {}
        self.edits_details = {}
        self.consolidated_cache = []
        self.consolidated_cache_valid = False

        self.current_store_index = 0
        self.current_selected_iid = None
        self.current_loaded_path = None
        self.current_filtered_ids = []
        self.visible_store_indices = []
        self.show_hidden_stores = bool(self.settings.get("show_hidden_stores", False))
        self._progress_reset_job = None

        self.is_fullscreen = False
        self.report_only_mode = False
        self._report_left_sash = None

        self.report_filter_mode_var = tk.StringVar(value="TODOS")
        self.report_filter_value_var = tk.StringVar(value="")
        self.report_filter_status_var = tk.StringVar(value="")
        self.report_filter_menu = None

        self.current_folder = ""
        self.known_store_files = set()
        self.new_store_poll_ms = 3000
        self._watch_new_files_job = None

        self.description_sort_mode = safe_str(self.settings.get("description_sort_mode", "")).upper()

        # Variável de pesquisa (sem barra fixa)
        self.search_var = tk.StringVar(value="")
        self.search_popup = None  # janela flutuante de pesquisa

        self._build_style()
        self._build_ui()
        self.apply_column_visibility()
        self.update_fullscreen_ui()
        self.update_report_view_ui()

        self.bind_all("<F11>", self.toggle_fullscreen)
        self.bind_all("<Escape>", self.on_escape_key)
        self.bind_all("<Control-p>", self.toggle_search_popup)  # Ctrl+P

        if (
            self.settings.get("reopen_last_folder_on_start")
            and self.settings.get("last_folder")
            and os.path.isdir(self.settings.get("last_folder"))
        ):
            self.after(150, lambda: self.open_folder(folder=self.settings.get("last_folder")))

    # -------------------- Janela flutuante de pesquisa --------------------
    def toggle_search_popup(self, event=None):
        """Abre ou fecha a janela flutuante de pesquisa."""
        if self.search_popup is not None and self.search_popup.winfo_exists():
            self.search_popup.destroy()
            self.search_popup = None
        else:
            self.create_search_popup()

    def create_search_popup(self):
        """Cria a janela flutuante de pesquisa, centralizada e arrastável."""
        popup = tk.Toplevel(self)
        popup.title("Pesquisar produto")
        popup.resizable(False, False)
        popup.overrideredirect(True)  # remove bordas para visual limpo
        popup.configure(bg="#2C3E50")

        # Frame interno com borda arredondada visual
        frame = tk.Frame(popup, bg="white", bd=1, relief="solid")
        frame.pack(fill="both", expand=True, padx=1, pady=2)

        # Campo de entrada
        entry = ttk.Entry(frame, width=35, font=("Segoe UI", 11))
        entry.pack(padx=10, pady=10, ipady=4)
        entry.focus_set()

        # Label de instrução
        lbl = ttk.Label(frame, text="Pequisar código ou descrição", font=("Segoe UI", 8))
        #lbl.pack(pady=(0, 8))

        # Tornar a janela arrastável
        def start_move(event):
            popup.x = event.x
            popup.y = event.y

        def do_move(event):
            x = popup.winfo_x() + (event.x - popup.x)
            y = popup.winfo_y() + (event.y - popup.y)
            popup.geometry(f"+{x}+{y}")

        popup.bind("<Button-1>", start_move)
        popup.bind("<B1-Motion>", do_move)

        # Atualizar a pesquisa em tempo real
        def on_search_change(event=None):
            self.search_var.set(entry.get().strip())
            self.refresh_table()

        entry.bind("<KeyRelease>", on_search_change)

        # Fechar com ESC ou Ctrl+P novamente
        def close_popup(event=None):
            self.search_popup = None
            popup.destroy()

        popup.bind("<Escape>", close_popup)
        popup.bind("<Control-p>", close_popup)

        # Posicionar no centro da tela principal
        popup.update_idletasks()
        w = popup.winfo_width()
        h = popup.winfo_height()
        x = self.winfo_x() + (self.winfo_width() // 2) - (w // 2)
        y = self.winfo_y() + (self.winfo_height() // 2) - (h // 2)
        popup.geometry(f"+{x}+{y}")

        self.search_popup = popup

        # Limpar variável de pesquisa se a janela for fechada sem digitar?
        # Não, mantém o último filtro. Mas ao fechar, pode limpar opcionalmente:
        # Se quiser limpar ao fechar, descomente:
        # popup.protocol("WM_DELETE_WINDOW", lambda: [self.search_var.set(""), self.refresh_table(), close_popup()])

    # -------------------- Fim da pesquisa flutuante --------------------

    def load_settings(self):
        data = DEFAULT_SETTINGS.copy()
        data["field_visibility"] = DEFAULT_FIELD_VISIBILITY.copy()

        if os.path.exists(self.config_path):
            try:
                with open(self.config_path, "r", encoding="utf-8") as f:
                    loaded = json.load(f)
                if isinstance(loaded, dict):
                    data.update(loaded)
                    vis = DEFAULT_FIELD_VISIBILITY.copy()
                    vis.update(loaded.get("field_visibility", {}))
                    data["field_visibility"] = vis
            except Exception:
                pass

        if "report_read_folder" not in data:
            data["report_read_folder"] = ""
        if "description_sort_mode" not in data:
            data["description_sort_mode"] = ""

        return data

    def save_settings(self):
        try:
            with open(self.config_path, "w", encoding="utf-8") as f:
                json.dump(self.settings, f, ensure_ascii=False, indent=2)
        except Exception as e:
            messagebox.showwarning(APP_TITLE, f"Não foi possível salvar as configurações.\n\n{e}")

    def get_field_visibility(self):
        vis = DEFAULT_FIELD_VISIBILITY.copy()
        vis.update(self.settings.get("field_visibility", {}))
        return vis

    def apply_column_visibility(self):
        vis = self.get_field_visibility()
        display_cols = [c for c in ALL_COLUMNS if vis.get(c, True)]
        if not display_cols:
            display_cols = ["descricao"]
        self.tree["displaycolumns"] = display_cols

    def center_window(self, win, master=None):
        win.update_idletasks()
        if master:
            x = master.winfo_x() + (master.winfo_width() - win.winfo_width()) // 2
            y = master.winfo_y() + (master.winfo_height() - win.winfo_height()) // 2
        else:
            x = (win.winfo_screenwidth() - win.winfo_width()) // 2
            y = (win.winfo_screenheight() - win.winfo_height()) // 2
        win.geometry(f"+{x}+{y}")

    def _build_style(self):
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except Exception:
            pass

        bg = "#F3F6FA"
        card = "#FFFFFF"
        dark = "#17212B"
        gray = "#667085"
        blue = "#2563EB"
        green = "#16A34A"

        self.configure(bg=bg)

        style.configure("App.TFrame", background=bg)
        style.configure("Card.TFrame", background=card)
        style.configure("HeaderTitle.TLabel", background=bg, foreground=dark, font=("Segoe UI", 15, "bold"))
        style.configure("HeaderSub.TLabel", background=bg, foreground=gray, font=("Segoe UI", 9))
        style.configure("CardTitle.TLabel", background=card, foreground=dark, font=("Segoe UI", 10, "bold"))
        style.configure("MetricTitle.TLabel", background=card, foreground=gray, font=("Segoe UI", 8))
        style.configure("MetricValue.TLabel", background=card, foreground=dark, font=("Segoe UI", 12, "bold"))
        style.configure("Primary.TButton", font=("Segoe UI", 9, "bold"))
        style.configure("Treeview", rowheight=26, font=("Segoe UI", 9), background="#FFFFFF", fieldbackground="#FFFFFF")
        style.configure("Treeview.Heading", font=("Segoe UI", 9, "bold"))
        style.map("Treeview", background=[("selected", "#DCEEFF")], foreground=[("selected", "#111827")])
        style.configure("Blue.TLabelframe", background=card)
        style.configure("Blue.TLabelframe.Label", background=card, foreground=blue, font=("Segoe UI", 10, "bold"))
        style.configure(
            "Green.Horizontal.TProgressbar",
            troughcolor="#E5E7EB",
            background=green,
            bordercolor="#E5E7EB",
            lightcolor=green,
            darkcolor=green,
        )

    def _build_ui(self):
        root = ttk.Frame(self, style="App.TFrame", padding=(10, 6, 10, 10))
        root.pack(fill="both", expand=True)
        self.root_container = root

        top = ttk.Frame(root, style="App.TFrame")
        top.pack(fill="x", pady=(0, 2))
        self.top_header = top

        left = ttk.Frame(top, style="App.TFrame")
        left.pack(side="left", fill="x", expand=True)

        ttk.Label(left, text=APP_TITLE, style="HeaderTitle.TLabel").pack(anchor="w")
        self.subtitle_var = tk.StringVar(value="Nenhuma pasta carregada")
        ttk.Label(left, textvariable=self.subtitle_var, style="HeaderSub.TLabel").pack(anchor="w", pady=(0, 0))

        right = ttk.Frame(top, style="App.TFrame")
        right.pack(side="right")

        self.btn_fullscreen = ttk.Button(right, text="Tela cheia", command=self.toggle_fullscreen, style="Primary.TButton")
        self.btn_fullscreen.pack(side="left", padx=2)

        self.btn_check_reports = ttk.Button(
            right,
            text="Ver PDF",
            command=self.verify_reports_folder,
            style="Primary.TButton",
        )
        self.btn_check_reports.pack(side="left", padx=2)

        self.btn_report_only = ttk.Button(right, text="Somente relatório", command=self.toggle_report_only_mode, style="Primary.TButton")
        self.btn_report_only.pack(side="left", padx=2)

        ttk.Button(right, text="Configurações", command=self.open_config_dialog, style="Primary.TButton").pack(side="left", padx=2)

        self.filter_var = tk.StringVar(value=self.settings.get("filter_mode", "TODOS"))

        self.progress_wrap = ttk.Frame(root, style="App.TFrame")
        self.progress_wrap.pack(fill="x", pady=(2, 4))

        self.progress_var = tk.DoubleVar(value=0)
        self.progress = ttk.Progressbar(
            self.progress_wrap,
            maximum=100,
            variable=self.progress_var,
            style="Green.Horizontal.TProgressbar",
        )
        self.progress.pack(side="left", fill="x", expand=True)

        self.progress_text_var = tk.StringVar(value="Pronto")
        self.progress_label = ttk.Label(self.progress_wrap, textvariable=self.progress_text_var, width=40)
        self.progress_label.pack(side="left", padx=(8, 0))

        metrics_wrap = ttk.Frame(root, style="App.TFrame")
        metrics_wrap.pack(fill="x", pady=METRICS_WRAP_PADY)
        self.metrics_wrap = metrics_wrap

        self.metric_loja = self._metric_card(metrics_wrap, "LOJA ATUAL")
        self.metric_progresso = self._metric_card(metrics_wrap, "PROGRESSO")
        self.metric_total_prod = self._metric_card(metrics_wrap, "TOTAL DE ITENS")
        self.metric_alterados_loja = self._metric_card(metrics_wrap, "ALTERADOS NESTA LOJA")
        self.metric_alterados_total = self._metric_card(metrics_wrap, "ALTERADOS NO TOTAL")

        self.metric_loja.pack(side="left", fill="x", expand=True, padx=(0, 4))
        self.metric_progresso.pack(side="left", fill="x", expand=True, padx=4)
        self.metric_total_prod.pack(side="left", fill="x", expand=True, padx=4)
        self.metric_alterados_loja.pack(side="left", fill="x", expand=True, padx=4)
        self.metric_alterados_total.pack(side="left", fill="x", expand=True, padx=(4, 0))

        center = ttk.PanedWindow(root, orient="horizontal")
        center.pack(fill="both", expand=True)
        self.center_pane = center

        left_panel = ttk.Frame(center, style="Card.TFrame", padding=6)
        center.add(left_panel, weight=1)
        self.left_panel = left_panel

        right_panel = ttk.Frame(center, style="Card.TFrame", padding=8)
        center.add(right_panel, weight=7)
        self.right_panel = right_panel

        top_lojas = ttk.Frame(left_panel, style="Card.TFrame")
        top_lojas.pack(fill="x", pady=(0, 4))

        ttk.Label(top_lojas, text="Lojas", style="CardTitle.TLabel").pack(side="left", anchor="w")

        self.btn_toggle_hidden = ttk.Button(
            top_lojas,
            text="Ver lojas editadas",
            command=self.toggle_show_hidden_stores,
        )
        self.btn_toggle_hidden.pack(side="right")

        self.store_listbox = tk.Listbox(
            left_panel,
            width=STORE_PANEL_WIDTH_CHARS,
            height=STORE_LIST_HEIGHT,
            font=("Segoe UI", 10),
            activestyle="none",
            borderwidth=0,
            highlightthickness=1,
            relief="solid",
            exportselection=False,
        )
        self.store_listbox.pack(fill="both", expand=True)
        self.store_listbox.bind("<<ListboxSelect>>", self.on_store_select)

        table_header = ttk.Frame(right_panel, style="Card.TFrame")
        table_header.pack(fill="x", pady=(0, 4))
        self.table_header = table_header

        header_left = ttk.Frame(table_header, style="Card.TFrame")
        header_left.pack(side="left", fill="x", expand=True)
        self.header_left = header_left

        self.table_title_var = tk.StringVar(value="Nenhuma loja carregada")
        self.table_title_label = ttk.Label(header_left, textvariable=self.table_title_var, font=("Segoe UI", 12, "bold"))
        self.table_title_label.pack(side="left", anchor="w")

        self.inline_progress_host = ttk.Frame(header_left, style="Card.TFrame")

        self.inline_progress_var = tk.DoubleVar(value=0)
        self.inline_progress = ttk.Progressbar(
            self.inline_progress_host,
            maximum=100,
            variable=self.inline_progress_var,
            style="Green.Horizontal.TProgressbar",
        )
        self.inline_progress.pack(side="left", fill="x", expand=True)

        self.inline_progress_text_var = tk.StringVar(value="Pronto")
        self.inline_progress_label = ttk.Label(
            self.inline_progress_host,
            textvariable=self.inline_progress_text_var,
            width=32,
        )
        self.inline_progress_label.pack(side="left", padx=(8, 0))

        header_right = ttk.Frame(table_header, style="Card.TFrame")
        header_right.pack(side="right")
        self.header_right = header_right

        # Barra de pesquisa removida - agora usa janela flutuante com Ctrl+P

        self.file_status_var = tk.StringVar(value="")
        self.file_status_label = ttk.Label(header_right, textvariable=self.file_status_var, foreground="#B45309")
        self.file_status_label.pack(side="right", padx=(10, 0))

        self.report_filter_status_label = ttk.Label(header_right, textvariable=self.report_filter_status_var, foreground="#667085")
        self.report_filter_status_label.pack(side="right")

        table_wrap = ttk.Frame(right_panel, style="Card.TFrame")
        table_wrap.pack(fill="both", expand=True)
        self.table_wrap = table_wrap

        self.tree = ttk.Treeview(table_wrap, columns=ALL_COLUMNS, show="headings", selectmode="browse")
        self.tree.pack(side="left", fill="both", expand=True)

        widths = {
            "codigo": 70,
            "descricao": 300,
            "um": 33,
            "at_venda": 74,
            "ult_compra": 76,
            "compra_sep": 76,
            "custo": 76,
            "ult_venda": 76,
            "venda_sep": 74,
            "preco_sugerido": 76,
            "margem": 76,
            "margem_padrao": 76,
            "dt_futura": 88,
            "novo_preco_editado": 106,
        }

        anchors = {
            "codigo": "center",
            "descricao": "w",
            "um": "center",
            "at_venda": "center",
            "ult_compra": "center",
            "compra_sep": "center",
            "custo": "center",
            "ult_venda": "center",
            "venda_sep": "center",
            "preco_sugerido": "center",
            "margem": "center",
            "margem_padrao": "center",
            "dt_futura": "center",
            "novo_preco_editado": "center",
        }

        for c in ALL_COLUMNS:
            self.tree.heading(c, text=COLUMN_HEADINGS[c])
            self.tree.column(c, width=widths[c], anchor=anchors[c], stretch=True)

        self.tree.heading("codigo", text=COLUMN_HEADINGS["codigo"], command=self.show_code_filter_menu_from_heading)
        self.tree.heading("descricao", text=self.get_descricao_heading_text(), command=self.toggle_description_sort)

        self.tree.tag_configure("changed", background="#86EFAC")
        self.tree.tag_configure("critical", background="#FECACA")
        self.tree.tag_configure("warning", background="#FFF4BF")

        ysb = ttk.Scrollbar(table_wrap, orient="vertical", command=self.tree.yview)
        xsb = ttk.Scrollbar(right_panel, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=ysb.set, xscrollcommand=xsb.set)
        ysb.pack(side="right", fill="y")
        xsb.pack(fill="x")
        self.xsb = xsb

        self.tree.bind("<<TreeviewSelect>>", self.on_row_select)
        self.tree.bind("<Double-1>", lambda e: self.focus_price_entry())
        self.tree.bind("<Return>", lambda e: self.focus_price_entry())

        editor_box = ttk.LabelFrame(
            right_panel,
            text="Edição do item selecionado",
            style="Blue.TLabelframe",
            padding=8,
        )
        editor_box.pack(fill="x", pady=(8, 0))
        self.editor_box = editor_box

        info_top = ttk.Frame(editor_box)
        info_top.pack(fill="x")

        self.selected_info_var = tk.StringVar(value="Nenhum item selecionado")
        ttk.Label(info_top, textvariable=self.selected_info_var, font=("Segoe UI", 9, "bold")).pack(side="left")

        info_mid = ttk.Frame(editor_box)
        info_mid.pack(fill="x", pady=(6, 6))

        self.info_codigo_var = tk.StringVar(value="Código: -")
        self.info_atual_var = tk.StringVar(value="Preço anterior: -")
        self.info_sug_var = tk.StringVar(value="Sugestão: -")
        self.info_margem_var = tk.StringVar(value="Margem: -")

        ttk.Label(info_mid, textvariable=self.info_codigo_var, width=18).pack(side="left")
        ttk.Label(info_mid, textvariable=self.info_atual_var, width=20).pack(side="left")
        ttk.Label(info_mid, textvariable=self.info_sug_var, width=18).pack(side="left")
        ttk.Label(info_mid, textvariable=self.info_margem_var, width=16).pack(side="left")

        edit_bar = ttk.Frame(editor_box)
        edit_bar.pack(fill="x")
        self.edit_bar = edit_bar

        ttk.Label(edit_bar, text="Preço alterado:", font=("Segoe UI", 9, "bold")).pack(side="left")
        self.new_price_var = tk.StringVar()
        self.new_price_entry = ttk.Entry(edit_bar, textvariable=self.new_price_var, width=18)
        self.new_price_entry.pack(side="left", padx=(6, 12))
        self.new_price_entry.bind("<Return>", self.on_enter_price)
        self.new_price_entry.bind("<FocusOut>", self.on_price_focus_out)

        ttk.Button(edit_bar, text="Aplicar", command=lambda: self.apply_current_edit(next_row=False)).pack(side="left")
        ttk.Button(edit_bar, text="Limpar", command=self.clear_current_edit).pack(side="left", padx=(6, 0))
        ttk.Button(edit_bar, text="Loja anterior", command=self.prev_store).pack(side="left", padx=(12, 0))
        ttk.Button(edit_bar, text="Próxima loja", command=self.next_store).pack(side="left", padx=(6, 0))
        ttk.Button(edit_bar, text="Gerar PDF", command=self.finalize_report).pack(side="left", padx=(6, 0))

    def _metric_card(self, parent, title):
        frame = ttk.Frame(parent, style="Card.TFrame", padding=METRIC_CARD_PADDING)
        ttk.Label(frame, text=title, style="MetricTitle.TLabel").pack(anchor="w")
        lbl_value = ttk.Label(frame, text="-", style="MetricValue.TLabel")
        lbl_value.pack(anchor="w", pady=(METRIC_VALUE_PADY, 0))
        frame.value_label = lbl_value
        return frame

    def open_config_dialog(self):
        ConfigDialog(self)

    def get_reports_search_folder(self):
        candidates = [
            safe_str(self.settings.get("report_read_folder")),
            safe_str(self.current_folder),
            safe_str(self.settings.get("last_folder")),
            safe_str(self.settings.get("post_pdf_target_folder")),
        ]
        for folder in candidates:
            if folder and os.path.isdir(folder):
                return folder
        return ""

    def choose_reports_search_folder(self):
        initial_dir = self.get_reports_search_folder() or os.getcwd()
        folder = filedialog.askdirectory(
            title="Selecione a pasta para verificar relatórios",
            initialdir=initial_dir,
        )
        if folder:
            self.settings["report_read_folder"] = folder
            self.save_settings()
            return folder
        return ""

    def open_path_in_system(self, path):
        try:
            if sys.platform.startswith("win"):
                os.startfile(path)
            elif sys.platform == "darwin":
                subprocess.Popen(["open", path])
            else:
                subprocess.Popen(["xdg-open", path])
            return True
        except Exception:
            return False

    def verify_reports_folder(self):
        folder = self.get_reports_search_folder()
        if not folder:
            folder = self.choose_reports_search_folder()
            if not folder:
                return

        try:
            pdfs = []
            for name in os.listdir(folder):
                full = os.path.join(folder, name)
                if os.path.isfile(full) and name.lower().endswith(".pdf"):
                    pdfs.append(full)

            if not pdfs:
                self.set_progress(0, "Nenhum relatório PDF encontrado")
                ask = messagebox.askyesno(
                    APP_TITLE,
                    "Nenhum relatório PDF foi encontrado nessa pasta.\n\nDeseja escolher outra pasta?",
                )
                if ask:
                    new_folder = self.choose_reports_search_folder()
                    if new_folder:
                        self.verify_reports_folder()
                return

            pdfs.sort(key=lambda p: os.path.getmtime(p), reverse=True)
            latest_pdf = pdfs[0]

            ok = self.open_path_in_system(latest_pdf)
            if ok:
                self.set_progress(100, f"Relatório aberto: {os.path.basename(latest_pdf)}")
            else:
                messagebox.showwarning(
                    APP_TITLE,
                    f"O relatório foi encontrado, mas não foi possível abrir automaticamente.\n\n{latest_pdf}",
                )
        except Exception as e:
            messagebox.showerror(APP_TITLE, f"Erro ao verificar a pasta de relatórios:\n\n{e}")

    def get_descricao_heading_text(self):
        base = COLUMN_HEADINGS["descricao"]
        if self.description_sort_mode == "AZ":
            return f"{base}  A-Z"
        if self.description_sort_mode == "ZA":
            return f"{base}  Z-A"
        return base

    def update_descricao_heading(self):
        try:
            self.tree.heading("descricao", text=self.get_descricao_heading_text(), command=self.toggle_description_sort)
        except Exception:
            pass

    def toggle_description_sort(self):
        if self.description_sort_mode == "AZ":
            self.description_sort_mode = "ZA"
        else:
            self.description_sort_mode = "AZ"
        self.settings["description_sort_mode"] = self.description_sort_mode
        self.save_settings()
        self.update_descricao_heading()
        self.refresh_table()

    def show_code_filter_menu_from_heading(self):
        menu = self.build_report_filter_menu()
        try:
            x = self.winfo_pointerx()
            y = self.winfo_pointery()
            menu.tk_popup(x, y)
        finally:
            try:
                menu.grab_release()
            except Exception:
                pass

    def toggle_fullscreen(self, event=None):
        self.set_fullscreen(not self.is_fullscreen)
        return "break"

    def exit_fullscreen(self, event=None):
        if self.is_fullscreen:
            self.set_fullscreen(False)
        return "break"

    def toggle_report_only_mode(self):
        self.set_report_only_mode(not self.report_only_mode)

    def _repack_root_layout(self):
        try:
            self.center_pane.pack_forget()
        except Exception:
            pass

        try:
            self.top_header.pack_forget()
        except Exception:
            pass

        try:
            self.progress_wrap.pack_forget()
        except Exception:
            pass

        try:
            self.metrics_wrap.pack_forget()
        except Exception:
            pass

        if not self.report_only_mode:
            self.top_header.pack(fill="x", pady=(0, 2))
            self.progress_wrap.pack(fill="x", pady=(2, 4))
            self.metrics_wrap.pack(fill="x", pady=METRICS_WRAP_PADY)

        self.center_pane.pack(fill="both", expand=True)

    def _hide_left_panel_from_pane(self):
        try:
            panes = [str(p) for p in self.center_pane.panes()]
            if str(self.left_panel) in panes:
                self.center_pane.forget(self.left_panel)
        except Exception:
            pass

    def _show_left_panel_in_pane(self):
        try:
            panes = [str(p) for p in self.center_pane.panes()]
            if str(self.left_panel) not in panes:
                try:
                    self.center_pane.insert(0, self.left_panel, weight=1)
                except Exception:
                    self.center_pane.add(self.left_panel, weight=1)
        except Exception:
            pass

    def _pack_right_panel_normal(self):
        try:
            self.table_header.pack_forget()
        except Exception:
            pass
        try:
            self.table_wrap.pack_forget()
        except Exception:
            pass
        try:
            self.xsb.pack_forget()
        except Exception:
            pass
        try:
            self.editor_box.pack_forget()
        except Exception:
            pass

        self.table_header.pack(fill="x", pady=(0, 4))
        self.table_wrap.pack(fill="both", expand=True)
        self.xsb.pack(fill="x")
        self.editor_box.pack(fill="x", pady=(8, 0))

    def _pack_right_panel_report_only(self):
        try:
            self.table_header.pack_forget()
        except Exception:
            pass
        try:
            self.table_wrap.pack_forget()
        except Exception:
            pass
        try:
            self.xsb.pack_forget()
        except Exception:
            pass
        try:
            self.editor_box.pack_forget()
        except Exception:
            pass

        self.table_header.pack(fill="x", pady=(0, 4))
        self.table_wrap.pack(fill="both", expand=True)
        self.xsb.pack(fill="x")
        self.editor_box.pack(fill="x", pady=(8, 0))

    def set_report_only_mode(self, enabled):
        enabled = bool(enabled)

        if enabled and not self.is_fullscreen:
            self.set_fullscreen(True)

        self.report_only_mode = enabled
        self.update_report_view_ui()
        self.after(50, self._refresh_layout_after_fullscreen)

    def update_report_view_ui(self):
        if self.report_only_mode:
            self.btn_report_only.config(text="Sair do relatório")

            try:
                self._report_left_sash = self.center_pane.sashpos(0)
            except Exception:
                pass

            self._hide_left_panel_from_pane()
            self._repack_root_layout()

            try:
                self.inline_progress_host.pack_forget()
            except Exception:
                pass
            self.inline_progress_host.pack(side="left", fill="x", expand=True, padx=(12, 0))

            self._pack_right_panel_report_only()
            self.table_title_label.configure(font=("Segoe UI", 13, "bold"))
        else:
            self.btn_report_only.config(text="Somente relatório")

            try:
                self.inline_progress_host.pack_forget()
            except Exception:
                pass

            self._show_left_panel_in_pane()
            self._repack_root_layout()
            self._pack_right_panel_normal()
            self.table_title_label.configure(font=("Segoe UI", 12, "bold"))

            try:
                if self._report_left_sash is not None:
                    self.center_pane.sashpos(0, self._report_left_sash)
            except Exception:
                pass

        self.update_fullscreen_ui()
        self.update_idletasks()

        if self.report_only_mode:
            try:
                total_w = self.center_pane.winfo_width()
                if total_w > 100:
                    self.center_pane.sashpos(0, 1)
            except Exception:
                pass

    def on_escape_key(self, event=None):
        if self.report_only_mode or self.is_fullscreen:
            if self.report_only_mode:
                self.set_report_only_mode(False)
            if self.is_fullscreen:
                self.set_fullscreen(False)
            self.update_idletasks()
            return "break"
        # Se a popup de pesquisa estiver aberta, fechar com Escape
        if self.search_popup and self.search_popup.winfo_exists():
            self.search_popup.destroy()
            self.search_popup = None
            return "break"
        return None

    def set_fullscreen(self, enabled):
        self.is_fullscreen = bool(enabled)
        try:
            self.attributes("-fullscreen", self.is_fullscreen)
        except Exception:
            if self.is_fullscreen:
                self.state("zoomed")
            else:
                try:
                    self.state("normal")
                    self.geometry("1600x860")
                except Exception:
                    pass

        if not self.is_fullscreen and self.report_only_mode:
            self.report_only_mode = False
            self.update_report_view_ui()

        self.update_fullscreen_ui()
        self.after(50, self._refresh_layout_after_fullscreen)

    def _refresh_layout_after_fullscreen(self):
        try:
            self.update_idletasks()
            self.tree.update_idletasks()
            self.tree.yview_moveto(0 if not self.tree.get_children() else self.tree.yview()[0])
        except Exception:
            pass

    def update_fullscreen_ui(self):
        meta = self.get_current_meta()
        if meta:
            loja_txt = safe_str(meta.get("loja")) or "-"
            title = f"Relatório da Loja {loja_txt}" if (self.report_only_mode or self.is_fullscreen) else f"Loja {loja_txt}"

            modo = self.report_filter_mode_var.get().strip()
            if modo != "TODOS":
                title += "  •  Filtro em CÓDIGO"

            if self.description_sort_mode == "AZ":
                title += "  •  Descrição A-Z"
            elif self.description_sort_mode == "ZA":
                title += "  •  Descrição Z-A"

            self.table_title_var.set(title)
        else:
            self.table_title_var.set("Nenhuma loja carregada")

        try:
            panes = [str(p) for p in self.center_pane.panes()]
            if str(self.left_panel) in panes:
                if self.is_fullscreen:
                    self.center_pane.paneconfig(self.left_panel, weight=0)
                    self.left_panel.pack_propagate(False)
                else:
                    self.center_pane.paneconfig(self.left_panel, weight=1)
        except Exception:
            pass

        if self.report_only_mode:
            self.btn_fullscreen.config(text="Tela cheia")
        else:
            self.btn_fullscreen.config(text="Sair da tela cheia" if self.is_fullscreen else "Tela cheia")

        self.update_descricao_heading()

    def build_report_filter_menu(self):
        if self.report_filter_menu is not None:
            try:
                self.report_filter_menu.destroy()
            except Exception:
                pass

        menu = tk.Menu(self, tearoff=0)

        menu.add_radiobutton(
            label="Todos",
            variable=self.report_filter_mode_var,
            value="TODOS",
            command=self.apply_report_filter_from_menu,
        )
        menu.add_radiobutton(
            label="Precisam de atenção",
            variable=self.report_filter_mode_var,
            value="PRECISAM DE ATENÇÃO",
            command=self.apply_report_filter_from_menu,
        )
        menu.add_radiobutton(
            label="Somente alterados",
            variable=self.report_filter_mode_var,
            value="SOMENTE ALTERADOS",
            command=self.apply_report_filter_from_menu,
        )

        menu.add_separator()

        menu.add_command(
            label="Margem abaixo de...",
            command=lambda: self.ask_report_filter_value("MARGEM ABAIXO DE"),
        )
        menu.add_command(
            label="Margem acima de...",
            command=lambda: self.ask_report_filter_value("MARGEM ACIMA DE"),
        )

        menu.add_separator()
        menu.add_command(label="Limpar filtro", command=self.clear_report_filter)

        self.report_filter_menu = menu
        return menu

    def ask_report_filter_value(self, mode):
        win = tk.Toplevel(self)
        win.title("Filtro do relatório")
        win.resizable(False, False)
        win.transient(self)
        win.grab_set()

        ttk.Frame(win, padding=12).pack(fill="both", expand=True)
        body = win.winfo_children()[0]

        texto = "Informe a margem para filtrar:"
        ttk.Label(body, text=texto).pack(anchor="w")

        value_var = tk.StringVar(value=self.report_filter_value_var.get())
        entry = ttk.Entry(body, textvariable=value_var, width=18)
        entry.pack(fill="x", pady=(8, 10))
        entry.focus_set()
        entry.selection_range(0, tk.END)

        buttons = ttk.Frame(body)
        buttons.pack(fill="x")

        def confirmar():
            raw = value_var.get().strip()
            val = money_to_float(raw)
            if val is None:
                messagebox.showwarning(APP_TITLE, "Informe um valor válido.", parent=win)
                entry.focus_set()
                return
            self.report_filter_mode_var.set(mode)
            self.report_filter_value_var.set(float_to_br(val))
            win.destroy()
            self.refresh_table()

        def cancelar():
            win.destroy()

        ttk.Button(buttons, text="Aplicar", command=confirmar).pack(side="left")
        ttk.Button(buttons, text="Cancelar", command=cancelar).pack(side="left", padx=(6, 0))

        win.bind("<Return>", lambda e: confirmar())
        win.bind("<Escape>", lambda e: cancelar())
        self.center_window(win, self)

    def apply_report_filter_from_menu(self):
        mode = self.report_filter_mode_var.get().strip()
        if mode in ("MARGEM ABAIXO DE", "MARGEM ACIMA DE"):
            raw = self.report_filter_value_var.get().strip()
            val = money_to_float(raw)
            if val is None:
                self.ask_report_filter_value(mode)
                return
            self.report_filter_value_var.set(float_to_br(val))
        else:
            self.report_filter_value_var.set("")
        self.refresh_table()

    def clear_report_filter(self):
        self.report_filter_mode_var.set("TODOS")
        self.report_filter_value_var.set("")
        self.report_filter_status_var.set("")
        self.refresh_table()

    def toggle_show_hidden_stores(self):
        self.show_hidden_stores = not self.show_hidden_stores
        self.settings["show_hidden_stores"] = self.show_hidden_stores
        self.save_settings()
        self.rebuild_store_listbox()
        self.select_current_store_in_listbox()

    def update_toggle_hidden_button_text(self):
        if self.show_hidden_stores:
            self.btn_toggle_hidden.config(text="Ocultar vistas")
        else:
            self.btn_toggle_hidden.config(text="Ver editadas")

    def set_progress(self, pct, text=""):
        pct = max(0, min(100, pct))
        self.progress_var.set(pct)
        self.inline_progress_var.set(pct)

        if text:
            self.progress_text_var.set(text)
            self.inline_progress_text_var.set(text)

        if self._progress_reset_job is not None:
            try:
                self.after_cancel(self._progress_reset_job)
            except Exception:
                pass
            self._progress_reset_job = None

        def reset_progress():
            self.progress_var.set(0)
            self.inline_progress_var.set(0)

        if pct >= 100:
            self._progress_reset_job = self.after(700, reset_progress)

        self.update_idletasks()

    def row_to_tags(self, row):
        tags = []
        marg = row.get("margem")
        marg_pad = row.get("margem_padrao")
        changed = row.get("novo_preco_editado") is not None

        if changed:
            tags.append("changed")
        elif marg is not None and marg < 0:
            tags.append("critical")
        elif marg is not None and marg_pad is not None and marg < marg_pad:
            tags.append("warning")

        return tuple(tags)

    def row_to_values(self, row):
        return (
            format_product_code(row["codigo"]),
            safe_str(row["descricao"]),
            safe_str(row["um"]),
            float_to_br(row["at_venda"]),
            float_to_br(row["ult_compra"]),
            maybe_number_to_br(row["compra_sep"]),
            float_to_br(row["custo"]),
            float_to_br(row["ult_venda"]),
            maybe_number_to_br(row["venda_sep"]),
            float_to_br(row["preco_sugerido"]),
            percent_to_br(row["margem"]),
            percent_to_br(row["margem_padrao"]),
            safe_str(row["dt_futura"]),
            float_to_br(row["novo_preco_editado"]),
        )

    def _clear_info_panel(self):
        self.info_codigo_var.set("Código: -")
        self.info_atual_var.set("Preço anterior: -")
        self.info_sug_var.set("Sugestão: -")
        self.info_margem_var.set("Margem: -")

    def update_store_dirty_status(self, path):
        meta = self.get_meta_by_path(path)
        if not meta:
            return

        changed_count = 0
        file_edits = self.edits_map.get(path, {})
        for _, v in file_edits.items():
            if v is not None:
                changed_count += 1

        meta["changed_count"] = changed_count
        meta["dirty"] = True if file_edits else False
        self.update_store_visibility(meta)

    def update_store_visibility(self, meta):
        meta["hidden"] = bool(meta.get("visited")) and int(meta.get("changed_count", 0) or 0) == 0

    def update_all_store_visibility(self):
        for meta in self.stores_meta:
            self.update_store_visibility(meta)

    def get_meta_by_path(self, path):
        for meta in self.stores_meta:
            if meta["arquivo"] == path:
                return meta
        return None

    def touch_cache(self, path):
        if path in self.loaded_cache:
            data = self.loaded_cache.pop(path)
            self.loaded_cache[path] = data

    def ensure_cache_limit(self):
        while len(self.loaded_cache) > CACHE_MAX_STORES:
            old_path, old_data = self.loaded_cache.popitem(last=False)
            if old_path == self.current_loaded_path:
                self.loaded_cache[old_path] = old_data
                break

    def get_current_meta(self):
        if not self.stores_meta:
            return None
        if self.current_store_index < 0 or self.current_store_index >= len(self.stores_meta):
            return None
        return self.stores_meta[self.current_store_index]

    def get_current_loaded_store(self):
        meta = self.get_current_meta()
        if not meta:
            return None
        path = meta["arquivo"]
        return self.loaded_cache.get(path)

    def get_next_visible_store_index(self, start_index=0):
        visible = [i for i, m in enumerate(self.stores_meta) if self.is_store_visible(m)]
        if not visible:
            return None

        for i in visible:
            if i >= start_index:
                return i
        return visible[0]

    def is_store_visible(self, meta):
        return self.show_hidden_stores or not meta.get("hidden", False)

    def stop_new_store_watcher(self):
        if self._watch_new_files_job is not None:
            try:
                self.after_cancel(self._watch_new_files_job)
            except Exception:
                pass
            self._watch_new_files_job = None

    def start_new_store_watcher(self):
        self.stop_new_store_watcher()
        self._watch_new_files_job = self.after(self.new_store_poll_ms, self.check_for_new_store_files)

    def check_for_new_store_files(self):
        self._watch_new_files_job = None

        try:
            if not self.current_folder or not os.path.isdir(self.current_folder):
                self.start_new_store_watcher()
                return

            current_files = list_excel_files(self.current_folder)
            current_files_set = set(current_files)
            new_files = [f for f in current_files if f not in self.known_store_files]

            if new_files:
                added_any = False
                errors = []

                for f in new_files:
                    try:
                        meta = scan_store_metadata(f)
                        meta["is_new"] = True
                        meta["visited"] = False
                        meta["hidden"] = False
                        meta["dirty"] = False
                        meta["loaded"] = False
                        meta["changed_count"] = int(meta.get("changed_count", 0) or 0)
                        self.stores_meta.append(meta)
                        added_any = True
                    except Exception as e:
                        errors.append(f"{os.path.basename(f)} -> {e}")

                self.known_store_files = current_files_set

                if added_any:
                    self.rebuild_store_listbox()
                    self.refresh_metrics()
                    self.set_progress(100, f"{len(new_files)} nova(s) loja(s) adicionada(s) à lista")

                if errors:
                    messagebox.showwarning(APP_TITLE, "Algumas novas lojas não puderam ser lidas:\n\n" + "\n".join(errors[:50]))
            else:
                self.known_store_files = current_files_set

        except Exception:
            pass

        self.start_new_store_watcher()

    def clear_store_list_after_post_process(self):
        self.stop_new_store_watcher()
        self.current_folder = ""
        self.known_store_files = set()

        self.stores_meta = []
        self.loaded_cache = OrderedDict()
        self.edits_map = {}
        self.edits_details = {}
        self.consolidated_cache = []
        self.consolidated_cache_valid = False

        self.current_store_index = 0
        self.current_selected_iid = None
        self.current_loaded_path = None
        self.current_filtered_ids = []
        self.visible_store_indices = []

        self.store_listbox.delete(0, tk.END)
        self.tree.delete(*self.tree.get_children())
        self.subtitle_var.set("Nenhuma pasta carregada")
        self.table_title_var.set("Nenhuma loja carregada")
        self.file_status_var.set("")
        self.selected_info_var.set("Nenhum item selecionado")
        self.new_price_var.set("")
        self.report_filter_status_var.set("")
        self.progress_text_var.set("Pronto")
        self.inline_progress_text_var.set("Pronto")
        self.progress_var.set(0)
        self.inline_progress_var.set(0)
        self._clear_info_panel()
        self.refresh_metrics()
        self.update_fullscreen_ui()

    def _release_loaded_file_handles(self):
        self.stop_new_store_watcher()

        for _, data in list(self.loaded_cache.items()):
            try:
                closer = getattr(data, "close", None)
                if callable(closer):
                    closer()
            except Exception:
                pass

            if isinstance(data, dict):
                for _, value in list(data.items()):
                    try:
                        closer = getattr(value, "close", None)
                        if callable(closer):
                            closer()
                    except Exception:
                        pass

        self.loaded_cache = OrderedDict()
        self.current_loaded_path = None

        try:
            self.update_idletasks()
        except Exception:
            pass

        gc.collect()

    def open_folder(self, folder=None):
        if not folder:
            initial_dir = self.settings.get("last_folder") or os.getcwd()
            folder = filedialog.askdirectory(
                title="Selecione a pasta com as planilhas",
                initialdir=initial_dir,
            )
        if not folder:
            return

        files = list_excel_files(folder)
        if not files:
            messagebox.showwarning(APP_TITLE, "Nenhum arquivo .xlsx ou .xls encontrado na pasta selecionada.")
            return

        self.stop_new_store_watcher()

        self.settings["last_folder"] = folder
        if not safe_str(self.settings.get("report_read_folder")):
            self.settings["report_read_folder"] = folder
        self.save_settings()

        self.set_progress(0, "Lendo lista de lojas...")
        self.subtitle_var.set(folder)
        self.stores_meta = []
        self.loaded_cache = OrderedDict()
        self.edits_map = {}
        self.edits_details = {}
        self.consolidated_cache = []
        self.consolidated_cache_valid = False
        self.current_store_index = 0
        self.current_selected_iid = None
        self.current_loaded_path = None
        self.current_filtered_ids = []
        self.visible_store_indices = []
        self.show_hidden_stores = bool(self.settings.get("show_hidden_stores", False))
        self.current_folder = folder
        self.known_store_files = set(files)

        self.store_listbox.delete(0, tk.END)
        self.tree.delete(*self.tree.get_children())
        self.table_title_var.set("Nenhuma loja carregada")
        self.file_status_var.set("")
        self.selected_info_var.set("Nenhum item selecionado")
        self.new_price_var.set("")
        self.report_filter_status_var.set("")
        self._clear_info_panel()

        errors = []
        total = len(files)

        for i, f in enumerate(files, start=1):
            try:
                meta = scan_store_metadata(f)
                meta["is_new"] = False
                self.stores_meta.append(meta)
            except Exception as e:
                errors.append(f"{os.path.basename(f)} -> {e}")

            pct = (i / total) * 100
            self.set_progress(pct, f"Lendo lista {i}/{total}")

        self.stores_meta.sort(
            key=lambda x: (
                int(x["loja"]) if safe_str(x["loja"]).isdigit() else 999999,
                safe_str(x["loja"]),
            )
        )
        self.update_all_store_visibility()
        self.rebuild_store_listbox()
        self.set_progress(100, f"Pasta carregada - {len(self.stores_meta)} loja(s)")

        if not self.stores_meta:
            messagebox.showerror(APP_TITLE, "Nenhuma planilha válida foi carregada.")
            return

        next_idx = self.get_next_visible_store_index(0)
        if next_idx is None:
            return

        self.current_store_index = next_idx
        self.select_current_store_in_listbox()
        self.load_selected_store(mark_visited=False)
        self.start_new_store_watcher()

        if errors:
            messagebox.showwarning(APP_TITLE, "Alguns arquivos não foram carregados:\n\n" + "\n".join(errors[:50]))

    def rebuild_store_listbox(self):
        current_path = None
        meta = self.get_current_meta()
        if meta:
            current_path = meta["arquivo"]

        self.store_listbox.delete(0, tk.END)
        self.visible_store_indices = []

        for real_idx, meta in enumerate(self.stores_meta):
            if not self.is_store_visible(meta):
                continue

            altered = meta.get("changed_count", 0)
            dirty_mark = "*" if meta.get("dirty") else ""
            loaded_mark = " •" if meta["arquivo"] in self.loaded_cache else ""
            visited_mark = " ✓" if meta.get("visited") else ""
            new_mark = " [NOVA]" if meta.get("is_new") else ""
            self.visible_store_indices.append(real_idx)
            self.store_listbox.insert(
                tk.END,
                f"Loja {meta['loja']} ({altered}) {dirty_mark}{loaded_mark}{visited_mark}{new_mark}".strip()
            )

        self.update_toggle_hidden_button_text()

        if current_path is not None:
            for list_idx, real_idx in enumerate(self.visible_store_indices):
                if self.stores_meta[real_idx]["arquivo"] == current_path:
                    self.store_listbox.selection_clear(0, tk.END)
                    self.store_listbox.selection_set(list_idx)
                    self.store_listbox.see(list_idx)
                    break

    def select_current_store_in_listbox(self):
        current_path = None
        meta = self.get_current_meta()
        if meta:
            current_path = meta["arquivo"]
        if current_path is None:
            return

        for list_idx, real_idx in enumerate(self.visible_store_indices):
            if self.stores_meta[real_idx]["arquivo"] == current_path:
                self.store_listbox.selection_clear(0, tk.END)
                self.store_listbox.selection_set(list_idx)
                self.store_listbox.see(list_idx)
                return

    def load_selected_store(self, mark_visited=False):
        meta = self.get_current_meta()
        if not meta:
            return

        if mark_visited:
            meta["visited"] = True
            self.update_store_visibility(meta)
            if meta.get("hidden", False) and not self.show_hidden_stores:
                next_idx = self.get_next_visible_store_index(self.current_store_index + 1)
                if next_idx is not None and next_idx != self.current_store_index:
                    self.current_store_index = next_idx
                    meta = self.get_current_meta()
                    if not meta:
                        return

        self.rebuild_store_listbox()
        self.select_current_store_in_listbox()

        path = meta["arquivo"]
        self.current_loaded_path = path
        meta["is_new"] = False

        if path in self.loaded_cache:
            self.touch_cache(path)
            self.refresh_table()
            self.set_progress(100, f"Loja {meta['loja']} pronta (cache)")
            self.rebuild_store_listbox()
            self.select_current_store_in_listbox()
            return

        try:
            self.set_progress(10, f"Carregando loja {meta['loja']}...")
            full_data = load_full_store_data(meta, self.edits_map)
            self.set_progress(75, f"Montando loja {meta['loja']}...")

            self.loaded_cache[path] = full_data
            meta["loaded"] = True
            meta["header_row"] = full_data["header_row"]
            meta["rows_count"] = full_data["rows_count"]
            meta["changed_count"] = full_data["changed_count"]
            meta["dirty"] = bool(self.edits_map.get(path))
            self.update_store_visibility(meta)

            self.ensure_cache_limit()
            self.refresh_table()
            self.rebuild_store_listbox()
            self.select_current_store_in_listbox()
            self.set_progress(100, f"Loja {meta['loja']} carregada")
        except Exception as e:
            meta["load_error"] = str(e)
            self.tree.delete(*self.tree.get_children())
            self.table_title_var.set(f"Loja {meta['loja']}")
            self.file_status_var.set("ERRO AO CARREGAR")
            self.set_progress(0, "Erro ao carregar loja")
            messagebox.showerror(
                APP_TITLE,
                f"Erro ao carregar a loja {meta['loja']}:\n\n{e}\n\n{traceback.format_exc()}",
            )

    def _filter_row_by_report_preferences(self, row):
        mode = self.report_filter_mode_var.get().strip()
        marg = row.get("margem")
        marg_pad = row.get("margem_padrao")
        changed = row.get("novo_preco_editado") is not None

        if mode == "TODOS":
            return True

        if mode == "PRECISAM DE ATENÇÃO":
            if marg is None:
                return False
            if marg < 0:
                return True
            if marg_pad is not None and marg < marg_pad:
                return True
            return False

        if mode == "SOMENTE ALTERADOS":
            return changed

        if mode in ("MARGEM ABAIXO DE", "MARGEM ACIMA DE"):
            ref = money_to_float(self.report_filter_value_var.get().strip())
            if ref is None or marg is None:
                return False
            if mode == "MARGEM ABAIXO DE":
                return marg < ref
            return marg > ref

        # Caso padrão (não deve ocorrer, mas garantia)
        return True

    def get_filtered_rows(self, store):
        rows = list(store["rows"])
        flt = self.filter_var.get()
        term = normalize_text(self.search_var.get())

        if flt == "SOMENTE ALTERADOS":
            rows = [r for r in rows if r.get("novo_preco_editado") is not None]
        elif flt == "MARGEM BAIXA/NEGATIVA":
            result = []
            for r in rows:
                marg = r.get("margem")
                marg_pad = r.get("margem_padrao")
                if marg is None:
                    continue
                if marg < 0:
                    result.append(r)
                elif marg_pad is not None and marg < marg_pad:
                    result.append(r)
            rows = result

        if term:
            filtrado = []
            for r in rows:
                hay = " ".join([safe_str(r.get("codigo")), safe_str(r.get("descricao"))])
                if term in normalize_text(hay):
                    filtrado.append(r)
            rows = filtrado

        rows = [r for r in rows if self._filter_row_by_report_preferences(r)]

        if self.description_sort_mode in ("AZ", "ZA"):
            reverse = self.description_sort_mode == "ZA"
            rows.sort(
                key=lambda r: (
                    normalize_text(safe_str(r.get("descricao"))),
                    safe_str(format_product_code(r.get("codigo"))),
                ),
                reverse=reverse,
            )

        mode = self.report_filter_mode_var.get().strip()
        sort_txt = ""
        if self.description_sort_mode == "AZ":
            sort_txt = " | Descrição A-Z"
        elif self.description_sort_mode == "ZA":
            sort_txt = " | Descrição Z-A"

        if mode == "TODOS":
            self.report_filter_status_var.set(f"{len(rows)} item(ns){sort_txt}")
        elif mode in ("MARGEM ABAIXO DE", "MARGEM ACIMA DE"):
            self.report_filter_status_var.set(
                f"{len(rows)} item(ns) | {mode.title()} {self.report_filter_value_var.get().strip()}{sort_txt}"
            )
        else:
            self.report_filter_status_var.set(f"{len(rows)} item(ns) | {mode.title()}{sort_txt}")

        self.current_filtered_ids = [int(r["source_index"]) for r in rows]
        return rows

    def refresh_metrics(self):
        meta = self.get_current_meta()
        if not meta:
            for card in [
                self.metric_loja,
                self.metric_progresso,
                self.metric_total_prod,
                self.metric_alterados_loja,
                self.metric_alterados_total,
            ]:
                card.value_label.config(text="-")
            return

        total_lojas = len(self.stores_meta)
        total_prod = meta.get("rows_count", 0)
        alt_loja = meta.get("changed_count", 0)
        alt_total = sum(m.get("changed_count", 0) for m in self.stores_meta)

        self.metric_loja.value_label.config(text=f"Loja {meta['loja']}")
        self.metric_progresso.value_label.config(text=f"{self.current_store_index + 1}/{total_lojas}")
        self.metric_total_prod.value_label.config(text=str(total_prod))
        self.metric_alterados_loja.value_label.config(text=str(alt_loja))
        self.metric_alterados_total.value_label.config(text=str(alt_total))

    def refresh_table(self):
        store = self.get_current_loaded_store()
        meta = self.get_current_meta()

        self.tree.delete(*self.tree.get_children())
        self.current_selected_iid = None
        self.new_price_var.set("")
        self.selected_info_var.set("Nenhum item selecionado")
        self._clear_info_panel()

        if not meta:
            self.table_title_var.set("Nenhuma loja carregada")
            self.file_status_var.set("")
            self.report_filter_status_var.set("")
            self.refresh_metrics()
            self.update_fullscreen_ui()
            return

        if self.report_only_mode:
            self.table_title_var.set(f"Relatório da Loja {meta['loja']}")
        elif self.is_fullscreen:
            self.table_title_var.set(f"Relatório da Loja {meta['loja']}")
        else:
            self.table_title_var.set(f"Loja {meta['loja']}")

        dirty_text = "ALTERAÇÕES EM MEMÓRIA" if meta.get("dirty") else "SOMENTE LEITURA"
        self.file_status_var.set(dirty_text)

        if not store:
            self.refresh_metrics()
            self.update_fullscreen_ui()
            return

        rows = self.get_filtered_rows(store)

        for r in rows:
            iid = f"{r['source_index']}"
            self.tree.insert("", "end", iid=iid, values=self.row_to_values(r), tags=self.row_to_tags(r))

        self.refresh_metrics()
        self.update_fullscreen_ui()

        children = self.tree.get_children()
        if children:
            self.tree.selection_set(children[0])
            self.tree.focus(children[0])
            self.tree.see(children[0])
            self.on_row_select()

    def on_store_select(self, event=None):
        selection = self.store_listbox.curselection()
        if not selection:
            return
        list_idx = int(selection[0])
        if list_idx < 0 or list_idx >= len(self.visible_store_indices):
            return
        real_idx = self.visible_store_indices[list_idx]
        if real_idx != self.current_store_index:
            self.current_store_index = real_idx
            self.load_selected_store(mark_visited=False)
        else:
            self.select_current_store_in_listbox()

    def prev_store(self):
        if not self.stores_meta:
            return

        current_meta = self.get_current_meta()
        if current_meta:
            current_meta["visited"] = True
            self.update_store_visibility(current_meta)

        visible_before = [i for i, m in enumerate(self.stores_meta) if self.is_store_visible(m)]
        if not visible_before:
            self.rebuild_store_listbox()
            self.tree.delete(*self.tree.get_children())
            self.table_title_var.set("Nenhuma loja pendente")
            self.file_status_var.set("")
            self.selected_info_var.set("Nenhum item selecionado")
            self.new_price_var.set("")
            self.report_filter_status_var.set("")
            self._clear_info_panel()
            self.refresh_metrics()
            self.update_fullscreen_ui()
            return

        prev_candidates = [i for i in visible_before if i < self.current_store_index]
        if prev_candidates:
            self.current_store_index = prev_candidates[-1]
        else:
            self.current_store_index = visible_before[-1]

        self.select_current_store_in_listbox()
        self.load_selected_store(mark_visited=False)

    def next_store(self):
        if not self.stores_meta:
            return

        current_meta = self.get_current_meta()
        if current_meta:
            current_meta["visited"] = True
            self.update_store_visibility(current_meta)

        visible_after = [i for i, m in enumerate(self.stores_meta) if self.is_store_visible(m)]
        if not visible_after:
            self.rebuild_store_listbox()
            self.tree.delete(*self.tree.get_children())
            self.table_title_var.set("Nenhuma loja pendente")
            self.file_status_var.set("")
            self.selected_info_var.set("Nenhum item selecionado")
            self.new_price_var.set("")
            self.report_filter_status_var.set("")
            self._clear_info_panel()
            self.refresh_metrics()
            self.update_fullscreen_ui()
            return

        next_candidates = [i for i in visible_after if i > self.current_store_index]
        if next_candidates:
            self.current_store_index = next_candidates[0]
        else:
            self.current_store_index = visible_after[0]

        self.select_current_store_in_listbox()
        self.load_selected_store(mark_visited=False)

    def get_current_row_ref(self):
        store = self.get_current_loaded_store()
        if not store:
            return None, None, None

        selection = self.tree.selection()
        if not selection:
            return store, None, None

        iid = selection[0]
        src_idx = int(iid)

        row = store["rows_map"].get(src_idx)
        return store, row, iid

    def on_row_select(self, event=None):
        _, row, iid = self.get_current_row_ref()
        self.current_selected_iid = iid
        if not row:
            self.selected_info_var.set("Nenhum item selecionado")
            self.new_price_var.set("")
            self._clear_info_panel()
            return

        self.new_price_var.set(float_to_br(row.get("novo_preco_editado")))
        self.selected_info_var.set(safe_str(row.get("descricao")))

        self.info_codigo_var.set(f"Código: {format_product_code(row.get('codigo')) or '-'}")
        self.info_atual_var.set(f"Preço anterior: {float_to_br(row.get('ult_venda')) or '-'}")
        self.info_sug_var.set(f"Sugestão: {float_to_br(row.get('preco_sugerido')) or '-'}")
        self.info_margem_var.set(f"Margem: {percent_to_br(row.get('margem')) or '-'}")

    def focus_price_entry(self):
        self.new_price_entry.focus_set()
        self.new_price_entry.selection_range(0, tk.END)

    def on_price_focus_out(self, event=None):
        txt = self.new_price_var.get().strip()
        if txt:
            formatted = sanitize_decimal_text_for_entry(txt)
            if formatted:
                self.new_price_var.set(formatted)

    def on_enter_price(self, event=None):
        self.apply_current_edit(next_row=True)
        return "break"

    def update_tree_row_visual(self, row):
        iid = f"{row['source_index']}"
        if not self.tree.exists(iid):
            return
        self.tree.item(iid, values=self.row_to_values(row), tags=self.row_to_tags(row))

    def update_edits_details(self, path, src_idx, novo, row, meta):
        if path not in self.edits_details:
            self.edits_details[path] = {}
        if novo is None:
            self.edits_details[path].pop(src_idx, None)
            if not self.edits_details[path]:
                del self.edits_details[path]
        else:
            self.edits_details[path][src_idx] = {
                "codigo": format_product_code(row.get("codigo", "")),
                "descricao": row.get("descricao", ""),
                "ult_venda": row.get("ult_venda"),
                "novo_preco": money_to_float(novo),
                "loja": safe_str(meta["loja"]),
                "source_idx": src_idx,
            }
        self.consolidated_cache_valid = False

    def apply_current_edit(self, next_row=False):
        store, row, iid = self.get_current_row_ref()
        meta = self.get_current_meta()
        if not row or not meta:
            messagebox.showwarning(APP_TITLE, "Selecione uma linha para editar.")
            return

        raw_price = self.new_price_var.get().strip()

        novo = None
        if raw_price:
            novo = money_to_float(raw_price)
            if novo is None:
                messagebox.showerror(APP_TITLE, "Preço alterado inválido.")
                return
            self.new_price_var.set(float_to_br(novo))

        path = meta["arquivo"]
        idx = int(row["source_index"])
        file_edits = self.edits_map.setdefault(path, {})

        row["novo_preco_editado"] = novo
        row["alterado"] = bool(novo is not None)

        if novo is None:
            file_edits[idx] = None
        else:
            file_edits[idx] = novo

        self.update_edits_details(path, idx, novo, row, meta)
        self.update_store_dirty_status(path)

        if path in self.loaded_cache:
            self.loaded_cache[path]["changed_count"] = meta["changed_count"]

        self.rebuild_store_listbox()
        self.select_current_store_in_listbox()
        self.update_tree_row_visual(row)
        self.file_status_var.set("ALTERAÇÕES EM MEMÓRIA")
        self.refresh_metrics()
        self.on_row_select()
        self.update_fullscreen_ui()

        if iid:
            try:
                self.tree.selection_set(iid)
                self.tree.focus(iid)
                self.tree.see(iid)
            except Exception:
                pass

        if next_row:
            self.select_next_row()

    def clear_current_edit(self):
        store, row, _ = self.get_current_row_ref()
        meta = self.get_current_meta()
        if not row or not meta:
            return

        path = meta["arquivo"]
        idx = int(row["source_index"])
        file_edits = self.edits_map.setdefault(path, {})

        row["novo_preco_editado"] = None
        row["alterado"] = False
        file_edits[idx] = None

        self.update_edits_details(path, idx, None, row, meta)
        self.update_store_dirty_status(path)

        if path in self.loaded_cache:
            self.loaded_cache[path]["changed_count"] = meta["changed_count"]

        self.new_price_var.set("")
        self.rebuild_store_listbox()
        self.select_current_store_in_listbox()
        self.update_tree_row_visual(row)
        self.file_status_var.set("ALTERAÇÕES EM MEMÓRIA")
        self.refresh_metrics()
        self.on_row_select()
        self.update_fullscreen_ui()

    def select_next_row(self):
        children = list(self.tree.get_children())
        selection = self.tree.selection()
        if not selection or not children:
            return

        cur = selection[0]
        try:
            idx = children.index(cur)
        except ValueError:
            return

        if idx < len(children) - 1:
            nxt = children[idx + 1]
            self.tree.selection_set(nxt)
            self.tree.focus(nxt)
            self.tree.see(nxt)
            self.on_row_select()
            self.focus_price_entry()
        else:
            self.next_store()
            self.focus_price_entry()

    def _build_current_store_export_rows(self, meta):
        path = meta["arquivo"]
        details = self.edits_details.get(path, {})
        rows = []

        for _, data in sorted(details.items(), key=lambda x: x[0]):
            novo_preco = money_to_float(data.get("novo_preco"))
            if novo_preco is None:
                continue

            rows.append({
                "LOJA": safe_str(data.get("loja")),
                "CODIGO": format_product_code(data.get("codigo")),
                "DESCRICAO": safe_str(data.get("descricao")),
                "PRECO_ANTERIOR": float_to_br(data.get("ult_venda")),
                "PRECO_ALTERADO": float_to_br(novo_preco),
            })

        return rows

    def save_current_store_file(self):
        meta = self.get_current_meta()
        if not meta:
            messagebox.showwarning(APP_TITLE, "Nenhuma loja carregada.")
            return

        rows = self._build_current_store_export_rows(meta)
        if not rows:
            messagebox.showinfo(APP_TITLE, "Não há alterações nesta loja para exportar.")
            return

        suggested = f"LOJA_{safe_str(meta['loja'])}_ALTERACOES.xlsx"
        path = filedialog.asksaveasfilename(
            title="Exportar alterações da loja",
            initialfile=suggested,
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
        )
        if not path:
            return

        try:
            self.set_progress(25, f"Exportando loja {meta['loja']}...")
            df = pd.DataFrame(rows)
            df = df.astype(str)

            with pd.ExcelWriter(path, engine="openpyxl") as writer:
                df.to_excel(writer, index=False, sheet_name="ALTERACOES_LOJA")

            self.set_progress(100, "Exportação concluída")
            messagebox.showinfo(APP_TITLE, f"Alterações da loja exportadas com sucesso:\n{path}")
        except Exception as e:
            messagebox.showerror(APP_TITLE, f"Erro ao exportar alterações da loja:\n{e}\n\n{traceback.format_exc()}")

    def save_all_files_internal(self, interactive=True):
        pending_count = sum(1 for m in self.stores_meta if m.get("dirty"))
        total_items = sum(m.get("changed_count", 0) for m in self.stores_meta)

        if interactive:
            if pending_count == 0:
                messagebox.showinfo(APP_TITLE, "Não há alterações em memória para exportar.")
            else:
                messagebox.showinfo(
                    APP_TITLE,
                    "As planilhas originais não são alteradas.\n\n"
                    f"Lojas com alterações em memória: {pending_count}\n"
                    f"Itens alterados em memória: {total_items}\n\n"
                    "Use:\n"
                    "- Exportar XLSX\n"
                    "- Exportar alterações da loja\n"
                    "- Gerar PDF"
                )

        return True, [], []

    def save_all_files(self):
        self.save_all_files_internal(interactive=True)

    def get_summary(self):
        lojas_total = len(self.stores_meta)
        lojas_com_alteracao = sum(1 for m in self.stores_meta if m.get("changed_count", 0) > 0)
        itens_alterados = sum(m.get("changed_count", 0) for m in self.stores_meta)
        groups = self.consolidate_changes_fast()

        return {
            "lojas_total": lojas_total,
            "lojas_com_alteracao": lojas_com_alteracao,
            "itens_alterados": itens_alterados,
            "agrupamentos": len(groups),
        }

    def consolidate_changes_fast(self):
        if self.consolidated_cache_valid:
            return self.consolidated_cache

        groups = {}
        for _, details in self.edits_details.items():
            for _, data in details.items():
                if data.get("novo_preco") is None:
                    continue

                codigo = format_product_code(data.get("codigo", ""))
                descricao = data.get("descricao", "")
                preco_novo = money_to_float(data.get("novo_preco"))
                loja = safe_str(data.get("loja", ""))
                ult_venda = money_to_float(data.get("ult_venda"))

                if preco_novo is None:
                    continue

                group_key = (codigo, normalize_text(descricao), f"{preco_novo:.2f}")
                if group_key not in groups:
                    groups[group_key] = {
                        "codigo": codigo or "SEM CÓDIGO",
                        "descricao": descricao,
                        "preco_anterior_por_loja": {},
                        "preco_alterado": preco_novo,
                        "lojas": [],
                        "qtd_lojas": 0,
                    }
                groups[group_key]["preco_anterior_por_loja"][loja] = ult_venda
                if loja not in groups[group_key]["lojas"]:
                    groups[group_key]["lojas"].append(loja)

        result = list(groups.values())
        for r in result:
            r["lojas"] = sorted(r["lojas"], key=lambda x: int(x) if x.isdigit() else x)
            r["qtd_lojas"] = len(r["lojas"])
            r["codigo"] = format_product_code(r.get("codigo", ""))

        result.sort(key=lambda x: (normalize_text(x["descricao"]), x["codigo"], x["preco_alterado"] or -999999))

        self.consolidated_cache = result
        self.consolidated_cache_valid = True
        return result

    def export_consolidated_xlsx(self):
        if not self.stores_meta:
            messagebox.showwarning(APP_TITLE, "Carregue uma pasta de planilhas primeiro.")
            return

        path = filedialog.asksaveasfilename(
            title="Salvar consolidado XLSX",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
        )
        if not path:
            return

        try:
            self.set_progress(20, "Gerando consolidado XLSX...")
            consolidated = self.consolidate_changes_fast()
            rows = []
            for group in consolidated:
                for loja in group["lojas"]:
                    rows.append({
                        "LOJA": safe_str(loja),
                        "CODIGO": format_product_code(group["codigo"]),
                        "DESCRICAO": safe_str(group["descricao"]),
                        "PRECO_ANTERIOR": float_to_br(group["preco_anterior_por_loja"].get(loja)),
                        "PRECO_ALTERADO": float_to_br(group["preco_alterado"]),
                    })
            df = pd.DataFrame(rows if rows else [{
                "LOJA": "",
                "CODIGO": "",
                "DESCRICAO": "SEM ALTERACOES",
                "PRECO_ANTERIOR": "",
                "PRECO_ALTERADO": "",
            }])

            df = df.astype(str)

            with pd.ExcelWriter(path, engine="openpyxl") as writer:
                df.to_excel(writer, index=False, sheet_name="CONSOLIDADO")

            self.set_progress(100, "Consolidado gerado")
            messagebox.showinfo(APP_TITLE, f"Arquivo salvo com sucesso:\n{path}")
        except Exception as e:
            messagebox.showerror(APP_TITLE, f"Erro ao salvar XLSX:\n{e}\n\n{traceback.format_exc()}")

    def get_loaded_source_files(self):
        files = []
        for meta in self.stores_meta:
            p = meta.get("arquivo")
            if p and os.path.exists(p) and p not in files:
                files.append(os.path.abspath(os.path.normpath(p)))
        return files

    def move_files_to_folder(self, files, target_folder):
        os.makedirs(target_folder, exist_ok=True)
        moved = []
        errors = []

        for src in files:
            try:
                src = os.path.abspath(os.path.normpath(src))
                base = os.path.basename(src)
                dst = os.path.join(target_folder, base)

                if os.path.abspath(src) == os.path.abspath(dst):
                    errors.append(f"{base} -> origem e destino são a mesma pasta")
                    continue

                if os.path.exists(dst):
                    stem, ext = os.path.splitext(base)
                    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    dst = os.path.join(target_folder, f"{stem}_{stamp}{ext}")

                shutil.move(src, dst)
                moved.append((src, dst))
            except Exception as e:
                errors.append(f"{os.path.basename(src)} -> {e}")

        return moved, errors

    def _send_to_recycle_bin_powershell(self, src):
        src = os.path.abspath(os.path.normpath(src))

        script = r'''
Add-Type -AssemblyName Microsoft.VisualBasic
$path = $args[0]
try {
    if ([System.IO.Directory]::Exists($path)) {
        [Microsoft.VisualBasic.FileIO.FileSystem]::DeleteDirectory(
            $path,
            [Microsoft.VisualBasic.FileIO.UIOption]::OnlyErrorDialogs,
            [Microsoft.VisualBasic.FileIO.RecycleOption]::SendToRecycleBin
        )
    }
    elseif ([System.IO.File]::Exists($path)) {
        [Microsoft.VisualBasic.FileIO.FileSystem]::DeleteFile(
            $path,
            [Microsoft.VisualBasic.FileIO.UIOption]::OnlyErrorDialogs,
            [Microsoft.VisualBasic.FileIO.RecycleOption]::SendToRecycleBin
        )
    }
    else {
        throw "Arquivo não encontrado: $path"
    }
}
catch {
    Write-Error $_.Exception.Message
    exit 1
}
'''
        result = subprocess.run(
            [
                "powershell",
                "-NoProfile",
                "-ExecutionPolicy",
                "Bypass",
                "-Command",
                script,
                src,
            ],
            capture_output=True,
            text=True,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
        if result.returncode != 0:
            err = (result.stderr or result.stdout or "").strip()
            raise RuntimeError(err or "Falha ao enviar arquivo para a lixeira via PowerShell.")

    def send_files_to_trash(self, files):
        moved = []
        errors = []

        for src in files:
            try:
                src = os.path.abspath(os.path.normpath(src))

                if not os.path.exists(src):
                    moved.append(src)
                    continue

                ok = False
                last_error = None

                if SEND2TRASH_OK and send2trash is not None:
                    try:
                        send2trash(src)
                        ok = True
                    except Exception as e:
                        last_error = e

                if not ok and sys.platform.startswith("win"):
                    try:
                        self._send_to_recycle_bin_powershell(src)
                        ok = True
                    except Exception as e:
                        last_error = e

                if not ok:
                    raise last_error if last_error else RuntimeError("Não foi possível enviar para a lixeira.")

                moved.append(src)

            except Exception as e:
                errors.append(f"{os.path.basename(src)} -> {e}")

        return moved, errors

    def post_pdf_process_files(self):
        files = self.get_loaded_source_files()
        if not files:
            return True, "Nenhum arquivo para processar após o PDF.", False

        move_enabled = bool(self.settings.get("post_pdf_move_enabled", False))
        trash_enabled = bool(self.settings.get("post_pdf_trash_enabled", False))
        target_folder = safe_str(self.settings.get("post_pdf_target_folder"))

        if move_enabled and trash_enabled:
            return False, "Configuração inválida: mover e lixeira estão ativos ao mesmo tempo.", False

        if not move_enabled and not trash_enabled:
            return True, "Nenhuma ação pós-PDF ativada.", False

        self._release_loaded_file_handles()

        if trash_enabled:
            moved, errors = self.send_files_to_trash(files)
            if errors:
                return False, "Alguns arquivos não foram enviados para a lixeira:\n\n" + "\n".join(errors[:50]), False
            return True, f"{len(moved)} arquivo(s) enviados para a lixeira.", True

        if move_enabled:
            if not target_folder:
                return False, "Pasta de destino não configurada.", False
            moved, errors = self.move_files_to_folder(files, target_folder)
            if errors:
                return False, "Alguns arquivos não foram movidos:\n\n" + "\n".join(errors[:50]), False
            return True, f"{len(moved)} arquivo(s) movidos para:\n{target_folder}", True

        return True, "Nenhuma ação executada.", False

    def finalize_report(self):
        if not self.stores_meta:
            messagebox.showwarning(APP_TITLE, "Carregue uma pasta de planilhas primeiro.")
            return

        if not self.users:
            messagebox.showerror(
                APP_TITLE,
                "Arquivo de autorização não encontrado ou vazio.\n\n"
                "Crie um dos arquivos ao lado do script:\n"
                "- autorizacao.json\n- autorizacao.xlsx\n- autorizacao.csv",
            )
            return

        dialog = FinalizeDialog(self, self.users, self.get_summary)
        self.wait_window(dialog)

        signer = dialog.result
        if not signer:
            return

        path = filedialog.asksaveasfilename(
            title="Salvar relatório PDF",
            defaultextension=".pdf",
            filetypes=[("PDF", "*.pdf")],
        )
        if not path:
            return

        try:
            self.set_progress(10, "Consolidando alterações...")
            consolidated = self.consolidate_changes_fast()
            resumo = self.get_summary()
            self.set_progress(70, "Gerando PDF...")
            build_pdf(path, consolidated, signer, resumo)
            self.set_progress(100, "PDF gerado")

            ok_post, msg_post, cleared_after_move = self.post_pdf_process_files()

            if ok_post and cleared_after_move:
                self.clear_store_list_after_post_process()

            if ok_post:
                messagebox.showinfo(APP_TITLE, f"PDF gerado com sucesso:\n{path}\n\n{msg_post}")
            else:
                messagebox.showwarning(
                    APP_TITLE,
                    f"PDF gerado com sucesso:\n{path}\n\nMas houve problema na ação pós-PDF:\n\n{msg_post}",
                )
        except Exception as e:
            messagebox.showerror(APP_TITLE, f"Erro ao gerar PDF:\n{e}\n\n{traceback.format_exc()}")

    def on_closing(self):
        self.stop_new_store_watcher()
        self.settings["filter_mode"] = self.filter_var.get()
        self.settings["search_text"] = self.search_var.get()
        self.settings["show_hidden_stores"] = self.show_hidden_stores
        self.settings["description_sort_mode"] = self.description_sort_mode
        self.save_settings()

        pendentes = sum(1 for m in self.stores_meta if m.get("dirty"))
        if pendentes > 0:
            r = messagebox.askyesnocancel(
                APP_TITLE,
                f"Existem {pendentes} loja(s) com alterações em memória.\n\nDeseja exportar o consolidado antes de sair?",
            )
            if r is None:
                return
            if r:
                self.export_consolidated_xlsx()
        self.destroy()


if __name__ == "__main__":
    app = PriceEditorApp()
    app.protocol("WM_DELETE_WINDOW", app.on_closing)
    app.mainloop()