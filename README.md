# Price

**Price** é um sistema desktop desenvolvido em Python para auxiliar no controle, análise e alteração de preços em operações de varejo, especialmente supermercados, lojas e centros de distribuição.

O objetivo do projeto é facilitar a conferência de listas de preços, análise de margens, identificação de produtos alterados e geração de relatórios consolidados em PDF.

---

## Objetivo do Projeto

O sistema foi criado para reduzir o trabalho manual no processo de alteração de preços, permitindo que o usuário carregue planilhas, visualize produtos, edite preços e gere um relatório final com as alterações realizadas.

Ele é especialmente útil para cenários onde existem várias lojas, grande volume de produtos e necessidade constante de atualização de preços com base em custos, compras, margens e notas fiscais.

---

## Principais Funcionalidades

- Leitura de planilhas Excel com dados de produtos e preços.
- Identificação automática de lojas e arquivos disponíveis.
- Interface gráfica desktop com foco em produtividade.
- Visualização de produtos em tabela.
- Edição de novos preços.
- Controle de produtos alterados.
- Cálculo e exibição de margens.
- Filtros para análise de produtos.
- Geração de relatório consolidado em PDF.
- Suporte a múltiplas lojas.
- Barra de progresso durante operações.
- Modo tela cheia.
- Modo somente relatório.
- Organização dos arquivos processados.
- Estrutura preparada para empacotamento como aplicativo Windows.

---

## Tecnologias Utilizadas

- Python
- Tkinter
- Pandas
- OpenPyXL
- ReportLab
- PyInstaller
- Inno Setup

---

## Estrutura Geral do Projeto

```text
Price/
├── app.ico
├── EULA.rtf
├── README.md
├── requirements.txt
├── price.py
├── instalador.iss
├── dist/
├── build/
├── output_installer/
└── arquivos de configuração locais
