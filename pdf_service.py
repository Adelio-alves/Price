# -*- coding: utf-8 -*-
"""
pdf_service.py
"""

from datetime import datetime
from html import escape

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

from constants import DATE_FMT, PDF_TITLE
from helpers import float_to_br, format_preco_anterior_resumo, format_product_code, safe_str

SYSTEM_VERSION = "v34.0 Steadfast Alfa 26"
SYSTEM_DEVELOPER = "Developer: Adelio Alves"
SITE_OFDEVER = "adelioalves.com"


def pdf_safe_text(value):
    text = safe_str(value)
    if not text:
        return ""
    return escape(text, quote=True)


def build_pdf(pdf_path, consolidated_rows, signer, resumo):
    styles = getSampleStyleSheet()
    style_normal = styles["Normal"]
    style_normal.fontName = "Helvetica"
    style_normal.fontSize = 8

    style_small = ParagraphStyle("small", parent=style_normal, fontSize=7, leading=8)
    style_small_center = ParagraphStyle(
        "small_center",
        parent=style_normal,
        fontSize=7,
        leading=8,
        alignment=TA_CENTER,
    )
    style_title = ParagraphStyle("title", parent=styles["Title"], fontName="Helvetica-Bold", fontSize=14, leading=16)
    style_sub = ParagraphStyle("sub", parent=style_normal, fontSize=9, leading=11)

    style_footer = ParagraphStyle(
        "footer",
        parent=style_normal,
        fontSize=9,
        leading=12,
        alignment=TA_CENTER,
    )

    style_bold = ParagraphStyle(
        "bold",
        parent=style_normal,
        fontName="Helvetica-Bold",
        fontSize=9,
        leading=12,
        alignment=TA_CENTER,
    )

    style_discreet_footer = ParagraphStyle(
        "discreet_footer",
        parent=style_normal,
        fontName="Helvetica",
        fontSize=6.5,
        leading=7.5,
        textColor=colors.HexColor("#6B7280"),
        alignment=TA_CENTER,
    )

    style_preco_alterado = ParagraphStyle(
        "preco_alterado",
        parent=style_normal,
        fontName="Helvetica-Bold",
        fontSize=7.3,
        leading=8.0,
        alignment=TA_CENTER,
    )

    style_preco_anterior = ParagraphStyle(
        "preco_anterior",
        parent=style_normal,
        fontName="Helvetica",
        fontSize=7.0,
        leading=8.0,
        alignment=TA_CENTER,
    )

    doc = SimpleDocTemplate(
        pdf_path,
        pagesize=landscape(A4),
        leftMargin=7 * mm,
        rightMargin=7 * mm,
        topMargin=7 * mm,
        bottomMargin=7 * mm,
        title=PDF_TITLE,
        author="Adelio Alves",
        subject="Relatório de precificação",
        creator="Sistema Steadfast v34",
        keywords="relatório, preços, lojas, alteração, precificação"
    )

    story = []
    generated_at = datetime.now().strftime(DATE_FMT)

    story.append(Paragraph(pdf_safe_text(PDF_TITLE), style_title))
    story.append(Spacer(1, 2.5 * mm))
    story.append(Paragraph(pdf_safe_text(f"Gerado em: {generated_at}"), style_sub))
    story.append(Paragraph(
        pdf_safe_text(
            f"Lojas revisadas: {resumo['lojas_total']} | "
            f"Lojas com alteração: {resumo['lojas_com_alteracao']} | "
            f"Itens alterados: {resumo['itens_alterados']} | "
            f"Agrupamentos finais: {resumo['agrupamentos']}"
        ),
        style_sub
    ))
    story.append(Spacer(1, 3 * mm))

    data = [[
        "CODIGO", "DESCRICAO", "PRECO ANTERIOR", "PRECO ALTERADO", "LOJAS", "QTD"
    ]]

    for r in consolidated_rows:
        lojas_txt = ", ".join([str(x) for x in r["lojas"]])
        preco_ant_txt = format_preco_anterior_resumo(r["preco_anterior_por_loja"], r["lojas"])
        codigo_pdf = format_product_code(r["codigo"])

        preco_anterior_paragraph = Paragraph(
            pdf_safe_text(preco_ant_txt) or "-",
            style_preco_anterior
        )

        preco_alterado_paragraph = Paragraph(
            pdf_safe_text(float_to_br(r["preco_alterado"]) or "-"),
            style_preco_alterado
        )

        data.append([
            pdf_safe_text(codigo_pdf) or "-",
            Paragraph(pdf_safe_text(r["descricao"]) or "-", style_small),
            preco_anterior_paragraph,
            preco_alterado_paragraph,
            Paragraph(pdf_safe_text(lojas_txt) or "-", style_small),
            pdf_safe_text(str(r["qtd_lojas"])),
        ])

    if len(data) == 1:
        data.append(["-", "SEM ALTERACOES", "-", "-", "-", "-"])

    table = Table(
        data,
        repeatRows=1,
        colWidths=[22 * mm, 75 * mm, 55 * mm, 36 * mm, 66 * mm, 13 * mm],
    )

    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#E8EEF7")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.HexColor("#1F2937")),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTNAME", (0, 1), (0, -1), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 7.3),
        ("LEADING", (0, 0), (-1, -1), 8.0),
        ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#C9D3E0")),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F9FBFD")]),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("ALIGN", (0, 0), (0, -1), "CENTER"),
        ("ALIGN", (2, 1), (3, -1), "CENTER"),
        ("ALIGN", (2, 0), (3, 0), "CENTER"),
        ("ALIGN", (5, 1), (5, -1), "CENTER"),
        ("LEFTPADDING", (0, 0), (-1, -1), 3),
        ("RIGHTPADDING", (0, 0), (-1, -1), 3),
        ("TOPPADDING", (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
    ]))

    story.append(table)
    story.append(Spacer(1, 4 * mm))
    story.append(Paragraph(pdf_safe_text("Fechamento do Relatório"), style_bold))
    story.append(Spacer(1, 1.5 * mm))

    fechamento_linha = (
        f"Responsável: <b>{pdf_safe_text(signer.get('nome'))}</b>"
        f" &nbsp;&nbsp;&nbsp;|&nbsp;&nbsp;&nbsp; "
        f"Função: <b>{pdf_safe_text(signer.get('funcao'))}</b>"
    )
    if safe_str(signer.get("setor")):
        fechamento_linha += (
            f" &nbsp;&nbsp;&nbsp;|&nbsp;&nbsp;&nbsp; "
            f"Setor: <b>{pdf_safe_text(signer.get('setor'))}</b>"
        )
    fechamento_linha += (
        f" &nbsp;&nbsp;&nbsp;|&nbsp;&nbsp;&nbsp; "
        f"Data/Hora: <b>{pdf_safe_text(generated_at)}</b>"
        f" &nbsp;&nbsp;&nbsp;|&nbsp;&nbsp;&nbsp; "
        f"Validação: <b>Autorizado por senha interna</b>"
    )

    story.append(Paragraph(fechamento_linha, style_footer))
    story.append(Spacer(1, 1.5 * mm))

    story.append(
        Paragraph(
            pdf_safe_text(f"Versão do sistema: {SYSTEM_VERSION} | {SYSTEM_DEVELOPER} | {SITE_OFDEVER}"),
            style_discreet_footer,
        )
    )
    story.append(Spacer(1, 2 * mm))

    aviso_texto = (
        "Atenção: esta é uma versão experimental do sistema e não deve ser usada para outro fim. "
        " É vedada a sua reprodução total ou parcial."
    )
    story.append(
        Paragraph(
            pdf_safe_text(aviso_texto),
            style_discreet_footer,
        )
    )

    doc.build(story)