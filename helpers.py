# -*- coding: utf-8 -*-
"""
helpers.py
"""

import math
import re
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP

import pandas as pd


def safe_str(v):
    if v is None:
        return ""
    try:
        if pd.isna(v):
            return ""
    except Exception:
        pass
    return str(v).strip()


def normalize_text(s):
    s = safe_str(s).upper()
    s = re.sub(r"\s+", " ", s)
    return s.strip()


def only_digits(s):
    s = re.sub(r"\D+", "", safe_str(s))
    return s.strip()


def format_product_code(v):
    if v is None:
        return ""

    try:
        if pd.isna(v):
            return ""
    except Exception:
        pass

    if isinstance(v, int):
        return str(v)

    if isinstance(v, float):
        try:
            if math.isnan(v):
                return ""
        except Exception:
            pass
        if v.is_integer():
            return str(int(v))
        s = f"{v}"
        if re.fullmatch(r"\d+\.0+", s):
            return s.split(".")[0]
        return s.strip()

    s = safe_str(v)
    if not s:
        return ""

    s = s.replace("\u00A0", " ").strip()

    if re.fullmatch(r"\d+", s):
        return s

    if re.fullmatch(r"\d+\.0+", s):
        return s.split(".")[0]

    if re.fullmatch(r"\d+,0+", s):
        return s.split(",")[0]

    return s


def _quantize_2(v):
    try:
        d = Decimal(str(v))
        return d.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    except Exception:
        return None


def _normalize_numeric_string(s):
    s = safe_str(s)
    if not s:
        return ""

    s = s.replace("\u00A0", " ")
    s = s.replace("R$", "").replace("r$", "")
    s = s.replace("%", "")
    s = re.sub(r"\s+", "", s)

    negative = False
    if s.startswith("(") and s.endswith(")"):
        negative = True
        s = s[1:-1]

    if s.startswith("-"):
        negative = True
        s = s[1:]

    s = re.sub(r"[^0-9,.\-]", "", s)

    if not s:
        return ""

    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "")
            s = s.replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s:
        parts = s.split(",")
        if len(parts) == 2:
            s = parts[0].replace(".", "") + "." + parts[1]
        else:
            s = s.replace(".", "").replace(",", ".")
    else:
        if s.count(".") > 1:
            parts = s.split(".")
            last = parts[-1]
            head = "".join(parts[:-1])
            s = head + "." + last

    if negative and s and not s.startswith("-"):
        s = "-" + s

    return s


def money_to_float(v):
    if v is None:
        return None
    if isinstance(v, bool):
        return None

    if isinstance(v, (int, float, Decimal)):
        try:
            if isinstance(v, float) and math.isnan(v):
                return None
        except Exception:
            pass
        q = _quantize_2(v)
        return float(q) if q is not None else None

    s = _normalize_numeric_string(v)
    if not s or s in ("-", ".", "-.", ","):
        return None

    try:
        q = Decimal(s).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
        return float(q)
    except (InvalidOperation, ValueError):
        return None


def float_to_br(v):
    num = money_to_float(v)
    if num is None:
        return ""
    q = _quantize_2(num)
    if q is None:
        return ""
    s = f"{q:,.2f}"
    s = s.replace(",", "X").replace(".", ",").replace("X", ".")
    return s


def percent_to_br(v):
    num = money_to_float(v)
    if num is None:
        return ""
    q = _quantize_2(num)
    if q is None:
        return ""
    s = f"{q:,.2f}"
    s = s.replace(",", "X").replace(".", ",").replace("X", ".")
    return s


def maybe_number_to_br(v):
    if v is None:
        return ""

    try:
        if pd.isna(v):
            return ""
    except Exception:
        pass

    if isinstance(v, (int, float, Decimal)):
        return float_to_br(v)

    s = safe_str(v)
    if not s:
        return ""

    s_clean = s.replace("\u00A0", " ").strip()

    if re.search(r"\d", s_clean):
        num = money_to_float(s_clean)
        if num is not None:
            return float_to_br(num)

    return s_clean


def sanitize_decimal_text_for_entry(text):
    return float_to_br(money_to_float(text))


def format_preco_anterior_resumo(preco_por_loja, lojas):
    valores = []
    for loja in lojas:
        valores.append(preco_por_loja.get(loja))

    valores_txt = [float_to_br(v) for v in valores if v is not None and float_to_br(v) != ""]
    valores_unicos = sorted(set(valores_txt))

    if len(valores_unicos) <= 1:
        return valores_unicos[0] if valores_unicos else ""

    partes = []
    for loja in lojas:
        partes.append(f"LJ {loja}: {float_to_br(preco_por_loja.get(loja))}")
    return " | ".join(partes)