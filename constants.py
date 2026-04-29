# -*- coding: utf-8 -*-
"""
constants.py
"""

import os

APP_TITLE = "Painel de Alteração de Preços"
PDF_TITLE = "RELATÓRIO CONSOLIDADO DE ALTERAÇÕES DE PREÇO"
DATE_FMT = "%d/%m/%Y %H:%M:%S"


def _ensure_app_data_dir():
    base_dir = os.getenv("LOCALAPPDATA")
    if not base_dir:
        base_dir = os.path.expanduser("~")
    app_dir = os.path.join(base_dir, "Price")
    os.makedirs(app_dir, exist_ok=True)
    return app_dir


APP_DATA_DIR = _ensure_app_data_dir()
CONFIG_FILE = os.path.join(APP_DATA_DIR, "painel_precos_config.json")

CACHE_MAX_STORES = 2
LISTING_SCAN_ROWS = 12
HEADER_SCAN_ROWS = 20

STORE_PANEL_WIDTH_CHARS = 10
STORE_LIST_HEIGHT = 24
METRIC_CARD_PADDING = 4
METRIC_VALUE_PADY = 0
METRICS_WRAP_PADY = (3, 4)

ALL_COLUMNS = (
    "codigo", "descricao", "um", "at_venda", "ult_compra", "compra_sep",
    "custo", "ult_venda", "venda_sep", "preco_sugerido",
    "margem", "margem_padrao", "dt_futura", "novo_preco_editado"
)

COLUMN_HEADINGS = {
    "codigo": "Código",
    "descricao": "Descrição",
    "um": "UM",
    "at_venda": "At. Venda",
    "ult_compra": "Ult. Compra",
    "compra_sep": "/ Compra",
    "custo": "Custo",
    "ult_venda": "Ult. Venda",
    "venda_sep": "/ Venda",
    "preco_sugerido": "Sugestão",
    "margem": "Marg %",
    "margem_padrao": "Marg. Pad.",
    "dt_futura": "Dt. Futura",
    "novo_preco_editado": "Preço Alterado",
}

DEFAULT_FIELD_VISIBILITY = {
    "codigo": True,
    "descricao": True,
    "um": True,
    "at_venda": True,
    "ult_compra": True,
    "compra_sep": True,
    "custo": True,
    "ult_venda": True,
    "venda_sep": True,
    "preco_sugerido": True,
    "margem": True,
    "margem_padrao": True,
    "dt_futura": True,
    "novo_preco_editado": True,
}

DEFAULT_SETTINGS = {
    "last_folder": "",
    "reopen_last_folder_on_start": True,
    "show_hidden_stores": False,
    "filter_mode": "TODOS",
    "search_text": "",
    "field_visibility": DEFAULT_FIELD_VISIBILITY.copy(),
    "post_pdf_move_enabled": False,
    "post_pdf_target_folder": "",
    "post_pdf_trash_enabled": False,
}