# -*- coding: utf-8 -*-
"""
creditos_app.py
Tela separada de créditos/orientações para o sistema
"""

import tkinter as tk
from tkinter import ttk
import webbrowser


APP_TITLE = "Créditos e Orientações"


class CreditosApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("420x620")
        self.minsize(860, 520)
        self.configure(bg="#F3F6FA")

        self._build_style()
        self._build_ui()

    def _build_style(self):
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except Exception:
            pass

        bg = "#F3F6FA"
        card = "#FFFFFF"
        title = "#17212B"
        text = "#334155"
        blue = "#2563EB"
        green = "#16A34A"
        border = "#D9E2EC"

        style.configure("App.TFrame", background=bg)
        style.configure("Card.TFrame", background=card, relief="flat")
        style.configure("Title.TLabel", background=bg, foreground=title, font=("Segoe UI", 18, "bold"))
        style.configure("SubTitle.TLabel", background=bg, foreground=text, font=("Segoe UI", 10))
        style.configure("SectionTitle.TLabel", background=card, foreground=title, font=("Segoe UI", 11, "bold"))
        style.configure("Body.TLabel", background=card, foreground=text, font=("Segoe UI", 10), justify="left")
        style.configure("Link.TButton", font=("Segoe UI", 9, "bold"))
        style.configure("Primary.TButton", font=("Segoe UI", 9, "bold"))
        style.configure("Footer.TLabel", background=bg, foreground="#64748B", font=("Segoe UI", 9))
        style.configure("Green.TLabel", background=card, foreground=green, font=("Segoe UI", 10, "bold"))

        self.colors = {
            "bg": bg,
            "card": card,
            "title": title,
            "text": text,
            "blue": blue,
            "green": green,
            "border": border,
        }

    def _build_ui(self):
        root = ttk.Frame(self, style="App.TFrame", padding=(14, 10, 14, 14))
        root.pack(fill="both", expand=True)

        header = ttk.Frame(root, style="App.TFrame")
        header.pack(fill="x", pady=(0, 10))

        ttk.Label(header, text="Créditos e Orientações", style="Title.TLabel").pack(anchor="w")
        ttk.Label(
            header,
            text="Janela separada para exibir créditos do sistema, fontes, links úteis, suporte e observações.",
            style="SubTitle.TLabel",
        ).pack(anchor="w", pady=(2, 0))

        main = ttk.Frame(root, style="App.TFrame")
        main.pack(fill="both", expand=True)

        canvas = tk.Canvas(
            main,
            bg=self.colors["bg"],
            highlightthickness=0,
            bd=0,
        )
        scrollbar = ttk.Scrollbar(main, orient="vertical", command=canvas.yview)
        self.content = ttk.Frame(canvas, style="App.TFrame")

        self.content.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=self.content, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        self._build_card_identificacao()
        self._build_card_sites()
        self._build_card_orientacoes()
        self._build_card_licencas()
        self._build_card_suporte()
        self._build_footer()

        self.bind_all("<MouseWheel>", lambda e: canvas.yview_scroll(int(-1 * (e.delta / 120)), "units"))

    def _make_card(self, parent):
        outer = tk.Frame(
            parent,
            bg=self.colors["border"],
            bd=0,
            highlightthickness=0
        )
        outer.pack(fill="x", pady=(0, 10))

        inner = ttk.Frame(outer, style="Card.TFrame", padding=12)
        inner.pack(fill="both", expand=True, padx=1, pady=1)
        return inner

    def _section_title(self, parent, text):
        ttk.Label(parent, text=text, style="SectionTitle.TLabel").pack(anchor="w", pady=(0, 8))

    def _body_label(self, parent, text):
        ttk.Label(parent, text=text, style="Body.TLabel", wraplength=860).pack(anchor="w")

    def _open_link(self, url):
        try:
            webbrowser.open(url)
        except Exception:
            pass

    def _add_link_row(self, parent, title, url, description=""):
        row = ttk.Frame(parent, style="Card.TFrame")
        row.pack(fill="x", pady=4)

        left = ttk.Frame(row, style="Card.TFrame")
        left.pack(side="left", fill="x", expand=True)

        ttk.Label(left, text=title, style="Green.TLabel").pack(anchor="w")
        ttk.Label(left, text=url, style="Body.TLabel").pack(anchor="w")
        if description:
            ttk.Label(left, text=description, style="Body.TLabel", wraplength=700).pack(anchor="w", pady=(2, 0))

        ttk.Button(
            row,
            text="Abrir site",
            style="Link.TButton",
            command=lambda u=url: self._open_link(u)
        ).pack(side="right", padx=(10, 0))

    def _build_card_identificacao(self):
        card = self._make_card(self.content)
        self._section_title(card, "1. Identificação do sistema")

        self._body_label(
            card,
            "Sistema: Painel de Alteração de Preços\n"
            "Finalidade: análise, conferência, edição e consolidação de preços por loja.\n"
            "Módulo atual: tela de créditos, referências e orientações gerais.\n"
            "Observação: este arquivo é independente do sistema principal e pode ser aberto separadamente."
        )

    def _build_card_sites(self):
        card = self._make_card(self.content)
        self._section_title(card, "2. Sites e referências")

        self._add_link_row(
            card,
            "Python",
            "https://www.python.org/",
            "Site oficial da linguagem usada para desenvolver o sistema."
        )
        self._add_link_row(
            card,
            "Pandas",
            "https://pandas.pydata.org/",
            "Biblioteca usada para tratamento de dados e exportações."
        )
        self._add_link_row(
            card,
            "OpenPyXL",
            "https://openpyxl.readthedocs.io/",
            "Biblioteca comum para leitura e escrita de arquivos Excel .xlsx."
        )
        self._add_link_row(
            card,
            "TkDocs / Tkinter",
            "https://tkdocs.com/",
            "Referência útil para componentes visuais usados na interface."
        )
        self._add_link_row(
            card,
            "Microsoft Excel",
            "https://www.microsoft.com/microsoft-365/excel",
            "Ferramenta de apoio para planilhas utilizadas no fluxo operacional."
        )

    def _build_card_orientacoes(self):
        card = self._make_card(self.content)
        self._section_title(card, "3. Orientações de uso")

        self._body_label(
            card,
            "• Sempre conferir a pasta selecionada antes de iniciar.\n"
            "• Validar se os preços digitados estão no formato esperado.\n"
            "• Revisar itens críticos, margens negativas e alterações feitas em memória.\n"
            "• Antes de mover ou enviar arquivos para a lixeira, garantir que o PDF já foi gerado corretamente.\n"
            "• Usar esta tela para registrar fontes, responsáveis, observações operacionais e links internos/externos."
        )

    def _build_card_licencas(self):
        card = self._make_card(self.content)
        self._section_title(card, "4. Créditos técnicos / bibliotecas")

        self._body_label(
            card,
            "Este sistema pode utilizar componentes e bibliotecas de terceiros conforme o projeto principal.\n\n"
            "Exemplos comuns:\n"
            "• Python\n"
            "• Tkinter / ttk\n"
            "• Pandas\n"
            "• OpenPyXL\n"
            "• send2trash\n\n"
            "Ajuste esta lista conforme os módulos realmente usados no seu ambiente."
        )

    def _build_card_suporte(self):
        card = self._make_card(self.content)
        self._section_title(card, "5. Responsável / suporte / observações")

        self._body_label(
            card,
            "Responsável: ______________________________\n"
            "Empresa/Setor: ____________________________\n"
            "Contato: _________________________________\n"
            "E-mail: __________________________________\n"
            "Versão interna: ___________________________\n"
            "Data da atualização: ______________________\n\n"
            "Observações:\n"
            "____________________________________________________________\n"
            "____________________________________________________________\n"
            "____________________________________________________________"
        )

        actions = ttk.Frame(card, style="Card.TFrame")
        actions.pack(fill="x", pady=(12, 0))

        ttk.Button(actions, text="Fechar", style="Primary.TButton", command=self.destroy).pack(side="left")

    def _build_footer(self):
        footer = ttk.Frame(self.content, style="App.TFrame")
        footer.pack(fill="x", pady=(4, 0))
        ttk.Label(
            footer,
            text="Arquivo independente para exibir créditos, referências e orientações do sistema.",
            style="Footer.TLabel"
        ).pack(anchor="center")


if __name__ == "__main__":
    app = CreditosApp()
    app.mainloop()