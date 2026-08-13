# eFatura App

Aplicação desktop (Python + Tkinter) para descarregar automaticamente os
Documentos Fiscais Eletrónicos (DFE) do portal [eFatura](https://pe.efatura.cv/)
da DNRE de Cabo Verde e gerar um mapa de controlo em Excel, pronto a usar
em contabilidade.

Dado um NIF, password e mês, a app:

1. Faz login no portal (via [Playwright](https://playwright.dev/)) e lista
   todos os documentos fiscais do mês.
2. Descarrega o XML e o PDF de cada documento (faturas, faturas-recibo,
   talões de venda, notas de crédito/débito, notas de devolução).
3. Gera um ficheiro Excel "Faturas controlo" por direção (Receção/compras e
   Emissão/vendas), com uma linha por documento: fornecedor/cliente, NIF,
   número, valores, IVA e retenção na fonte.

Consulta o [LEIA-ME.txt](LEIA-ME.txt) para instruções de uso e instalação.

## Estrutura

- `eFatura_App.py` — interface gráfica (Tkinter) e orquestração do fluxo.
- `efatura_downloader.py` — login no portal e chamadas à API do eFatura.
- `gerar_mapa_controlo.py` — leitura dos XML descarregados para o mapa Excel.

## Aviso

Esta ferramenta não submete nem importa nada automaticamente em nenhum
sistema — produz sempre ficheiros de trabalho (Excel, XML, PDF) para
revisão humana antes de qualquer uso contabilístico/fiscal.
