# -*- coding: utf-8 -*-
"""
auth_service.py
"""

import json
import os

import pandas as pd

from helpers import normalize_text, safe_str


def build_column_map(df):
    mapping = {}
    for c in df.columns:
        mapping[normalize_text(c)] = c
    return mapping


def pick_col(colmap, *names):
    for n in names:
        if n in colmap:
            return colmap[n]
    return None


def load_authorization_file(base_dir):
    json_path = os.path.join(base_dir, "autorizacao.json")
    xlsx_path = os.path.join(base_dir, "autorizacao.xlsx")
    csv_path = os.path.join(base_dir, "autorizacao.csv")

    users = []

    if os.path.exists(json_path):
        with open(json_path, "r", encoding="utf-8") as f:
            data = json.load(f)

        for r in data:
            users.append({
                "senha": safe_str(r.get("senha")),
                "nome": safe_str(r.get("nome")),
                "funcao": safe_str(r.get("funcao")),
                "setor": safe_str(r.get("setor")),
                "ativo": bool(r.get("ativo", True)),
            })

    elif os.path.exists(xlsx_path):
        df = pd.read_excel(xlsx_path, engine="openpyxl")
        colmap = build_column_map(df)

        senha_col = pick_col(colmap, "SENHA")
        nome_col = pick_col(colmap, "NOME")
        funcao_col = pick_col(colmap, "FUNCAO", "FUNÇÃO")
        setor_col = pick_col(colmap, "SETOR")
        ativo_col = pick_col(colmap, "ATIVO")

        for _, r in df.iterrows():
            users.append({
                "senha": safe_str(r[senha_col]) if senha_col else "",
                "nome": safe_str(r[nome_col]) if nome_col else "",
                "funcao": safe_str(r[funcao_col]) if funcao_col else "",
                "setor": safe_str(r[setor_col]) if setor_col else "",
                "ativo": safe_str(r[ativo_col]).upper() not in ("", "0", "NÃO", "NAO", "NO", "FALSE", "INATIVO") if ativo_col else True,
            })

    elif os.path.exists(csv_path):
        df = pd.read_csv(csv_path, sep=None, engine="python")
        colmap = build_column_map(df)

        senha_col = pick_col(colmap, "SENHA")
        nome_col = pick_col(colmap, "NOME")
        funcao_col = pick_col(colmap, "FUNCAO", "FUNÇÃO")
        setor_col = pick_col(colmap, "SETOR")
        ativo_col = pick_col(colmap, "ATIVO")

        for _, r in df.iterrows():
            users.append({
                "senha": safe_str(r[senha_col]) if senha_col else "",
                "nome": safe_str(r[nome_col]) if nome_col else "",
                "funcao": safe_str(r[funcao_col]) if funcao_col else "",
                "setor": safe_str(r[setor_col]) if setor_col else "",
                "ativo": safe_str(r[ativo_col]).upper() not in ("", "0", "NÃO", "NAO", "NO", "FALSE", "INATIVO") if ativo_col else True,
            })

    return users


def find_user_by_password(users, senha):
    senha = safe_str(senha)
    for u in users:
        if safe_str(u.get("senha")) == senha and bool(u.get("ativo", True)):
            return u
    return None