"""
eFatura App - Download e preenchimento automatico do mapa de controlo.

Cola o NIF e a password do portal eFatura da DNRE (Cabo Verde), escolhe o
mes e a(s) direcao(oes) desejada(s) (Emissao/Rececao), carrega em
"Processar" e o app:
  1. Faz login no portal e lista todos os documentos do mes (a API nao
     permite filtrar por direcao na listagem - isso so' se sabe por
     documento), mas so' descarrega o XML dos documentos cuja direcao
     esteja marcada, poupando tempo quando so' se quer uma delas.
  2. Lista as faturas/faturas-recibo do mes num ficheiro Excel por
     direcao "Faturas controlo {empresa} {mes} - Recebidas/Emitidas.xlsx"
     na pasta Saida/, no mesmo formato usado pela contabilidade (sem
     classificacao contabilistica - isso e feito manualmente depois).
"""
import logging
import os
import queue
import re
import sys
import threading
import tkinter as tk
from collections import Counter
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from pathlib import Path
from tkinter import messagebox, scrolledtext, ttk

import pandas as pd

if getattr(sys, "frozen", False):
    APP_DIR = Path(sys.executable).resolve().parent
else:
    APP_DIR = Path(__file__).resolve().parent

# Usa o Chromium empacotado junto ao app (pasta playwright-browsers/) em vez
# de depender de uma instalacao previa do Playwright na maquina do colega.
_BROWSERS_DIR = APP_DIR / "playwright-browsers"
if _BROWSERS_DIR.exists():
    os.environ["PLAYWRIGHT_BROWSERS_PATH"] = str(_BROWSERS_DIR)

sys.path.insert(0, str(APP_DIR))

from efatura_downloader import (  # noqa: E402
    descarregar_pdf,
    descarregar_xml,
    listar_documentos,
    login_e_capturar_sessao,
    mes_para_intervalo,
    sanitizar_nome,
)
from gerar_mapa_controlo import carregar_documentos_de_pasta, mes_nome_pt  # noqa: E402

SAIDA_DIR = APP_DIR / "Saida"
DOWNLOAD_MAX_WORKERS = 8

# Tipos de documento a incluir no mapa - confirmados via tag raiz do XML real:
#   1 <Invoice> FT (Fatura), 2 <Invoice> FR (Fatura-Recibo),
#   3 <SalesReceipt> TV (Talao de Venda), 5 <CreditNote> NC (Nota de Credito),
#   8 <ReturnNote> DV (Nota de Devolucao).
# O tipo 6 (ND, Nota de Debito) nao apareceu em nenhuma amostra real
# testada - mantido por inferencia logica da sequencia.
# Excluidos: 4 (Recibo, so confirma pagamento de fatura ja emitida) e
# 7 <Transport> (Documento de Transporte, sem valor fiscal de venda/compra).
PREFIXO_TIPO = {
    "1": "FTE",
    "2": "FRE",
    "3": "TVE",
    "5": "NCE",
    "6": "NDE",
    "8": "DVE",
}
TIPOS_PROCESSAVEIS = set(PREFIXO_TIPO.keys())


def caminho_longo(caminho: Path) -> Path:
    """No Windows, caminhos com mais de 260 caracteres falham (FileNotFoundError)
    a nao ser que se use o prefixo de caminho longo \\\\?\\ - isto acontece
    facilmente quando o nome da empresa/fornecedor sao compridos e o colega
    instalou a app numa pasta mais funda que a nossa (ex: Documents\\dist\\Saida).
    O prefixo evita o limite sem precisar de alterar nada na maquina do
    utilizador (registo, direitos de admin, etc.)."""
    if os.name != "nt":
        return caminho
    resolvido = str(caminho.resolve())
    if resolvido.startswith("\\\\?\\"):
        return caminho
    return Path("\\\\?\\" + resolvido)


class GuiLogHandler(logging.Handler):
    def __init__(self, fila: queue.Queue):
        super().__init__()
        self.fila = fila

    def emit(self, record):
        self.fila.put(self.format(record))


def criar_logger(fila: queue.Queue) -> logging.Logger:
    logger = logging.getLogger(f"efatura_app_{id(fila)}")
    logger.setLevel(logging.INFO)
    logger.handlers.clear()
    handler = GuiLogHandler(fila)
    handler.setFormatter(logging.Formatter("%(message)s"))
    logger.addHandler(handler)
    return logger


def _direcao_doc(doc: dict, nif_normalizado: str) -> tuple[str, bool]:
    """Devolve (direcao, teve_fallback). direcao e' 'E' ou 'R'."""
    direcao = doc.get("IssueDirection")
    if direcao in ("E", "R"):
        return direcao, False
    # Fallback caso o campo venha em falta/inesperado: deduz a direcao
    # comparando o NIF do emissor com o NIF autenticado.
    emitter_tax_id = str(doc.get("EmitterTaxId", "")).strip()
    return ("E" if emitter_tax_id == nif_normalizado else "R"), True


def descarregar_documentos(
    nif: str, password: str, mes: str, incluir_emissao: bool, incluir_rececao: bool,
    logger: logging.Logger, progresso_callback=None,
):
    """progresso_callback(fase, concluidos, total), se indicado, e' chamado
    com fase="listagem" (paginas) e depois fase="download" (documentos)."""
    data_inicio, data_fim = mes_para_intervalo(mes)
    logger.info("A autenticar no portal eFatura (NIF %s)...", nif)
    headers_auth, nome_empresa = login_e_capturar_sessao(nif, password, logger)
    logger.info("Login OK. Empresa: %s", nome_empresa)

    documentos = listar_documentos(
        headers_auth, data_inicio, data_fim, logger,
        progress_callback=(lambda c, t: progresso_callback("listagem", c, t)) if progresso_callback else None,
    )
    logger.info("%s documentos encontrados para %s.", len(documentos), mes)

    nif_normalizado = str(nif).strip()
    contagem_direcao: Counter = Counter()
    contagem_tipo: Counter = Counter()
    fallback_count = 0
    documentos_a_descarregar = []
    for doc in documentos:
        direcao, teve_fallback = _direcao_doc(doc, nif_normalizado)
        contagem_direcao[direcao] += 1
        if teve_fallback:
            fallback_count += 1

        tipo = str(doc.get("DocumentTypeCode", ""))
        contagem_tipo[tipo] += 1

        direcao_ok = (direcao == "E" and incluir_emissao) or (direcao == "R" and incluir_rececao)
        if direcao_ok and tipo in TIPOS_PROCESSAVEIS:
            documentos_a_descarregar.append((doc, direcao))

    logger.info(
        "Classificacao por direcao (todo o mes): Emissao=%s, Rececao=%s (campo em falta/invalido em %s documentos)",
        contagem_direcao["E"], contagem_direcao["R"], fallback_count,
    )
    logger.info(
        "Classificacao por tipo (todo o mes): %s",
        ", ".join(f"{PREFIXO_TIPO.get(t, f'Tipo{t}')}={n}" for t, n in sorted(contagem_tipo.items())),
    )
    ignorados = len(documentos) - len(documentos_a_descarregar)
    if ignorados:
        logger.info(
            "%s documentos fora dos filtros de direcao/tipo selecionados - nao vao ser descarregados.", ignorados,
        )

    destino_mes = SAIDA_DIR / "01_Documentos_fiscais" / sanitizar_nome(nome_empresa) / mes
    destino_emissao = destino_mes / "Emissão"
    destino_rececao = destino_mes / "Receção"
    destino_emissao_pdf = destino_emissao / "PDF"
    destino_rececao_pdf = destino_rececao / "PDF"

    # So' cria/limpa a pasta da(s) direcao(oes) que vai(vao) ser mesmo
    # descarregada(s) - evita trabalho e confusao sobre pastas nao usadas.
    # O PDF fica numa subpasta para o glob("*.xml") do gerador do mapa
    # (nao recursivo) nunca a varrer.
    if incluir_emissao:
        caminho_longo(destino_emissao).mkdir(parents=True, exist_ok=True)
        caminho_longo(destino_emissao_pdf).mkdir(parents=True, exist_ok=True)
        for antigo in caminho_longo(destino_emissao).glob("*.xml"):
            antigo.unlink(missing_ok=True)
        for antigo in caminho_longo(destino_emissao_pdf).glob("*.pdf"):
            antigo.unlink(missing_ok=True)
    if incluir_rececao:
        caminho_longo(destino_rececao).mkdir(parents=True, exist_ok=True)
        caminho_longo(destino_rececao_pdf).mkdir(parents=True, exist_ok=True)
        for antigo in caminho_longo(destino_rececao).glob("*.xml"):
            antigo.unlink(missing_ok=True)
        for antigo in caminho_longo(destino_rececao_pdf).glob("*.pdf"):
            antigo.unlink(missing_ok=True)

    def _nome_descritivo_pdf(doc: dict, direcao: str) -> str:
        """Ex: "FTE1361 EMPRESA_EXEMPLO_LDA 01-07-2026" - tipo+numero, contraparte
        (fornecedor na Receção, cliente na Emissão) e data (dd-mm-aaaa)."""
        tipo_codigo = str(doc.get("DocumentTypeCode", ""))
        prefixo = PREFIXO_TIPO.get(tipo_codigo, f"Tipo{tipo_codigo}")
        numero = doc.get("DocumentNumber", "")
        contraparte = str((doc.get("ReceiverName") if direcao == "E" else doc.get("EmitterName")) or "").strip()

        data_str = str(doc.get("AuthorizedDateTime") or doc.get("IssueDateTime") or "")
        data_parte = data_str.split(" ", 1)[0]
        try:
            ano, mes, dia = data_parte.split("-")
            data_fmt = f"{dia}-{mes}-{ano}"
        except ValueError:
            data_fmt = data_parte

        return sanitizar_nome(f"{prefixo}{numero} {contraparte} {data_fmt}")

    def _baixar_um(doc: dict, direcao: str) -> bool:
        iud = doc.get("Id")
        if not iud:
            logger.warning("Documento sem Id (IUD), ignorado: %s", doc)
            return False

        tipo = sanitizar_nome(f"Tipo{doc.get('DocumentTypeCode', 'Desconhecido')}")
        emissor = sanitizar_nome(str(doc.get("EmitterName", "Emissor")))
        data_doc = sanitizar_nome(str(doc.get("AuthorizedDateTime") or doc.get("IssueDateTime") or ""))
        destino = destino_emissao if direcao == "E" else destino_rececao
        destino_pdf = destino_emissao_pdf if direcao == "E" else destino_rececao_pdf

        base_nome = f"{data_doc}_{tipo}_{emissor}_{iud}"
        try:
            conteudo_xml = descarregar_xml(headers_auth, iud, logger)
            caminho_longo(destino / f"{base_nome}.xml").write_bytes(conteudo_xml)
        except Exception as exc:
            logger.error("Falha ao descarregar XML do IUD %s: %s", iud, exc)
            return False

        try:
            conteudo_pdf = descarregar_pdf(headers_auth, iud, logger)
            nome_pdf = _nome_descritivo_pdf(doc, direcao)
            caminho_longo(destino_pdf / f"{nome_pdf}.pdf").write_bytes(conteudo_pdf)
        except Exception as exc:
            # PDF e' um extra - falhar aqui nao deve tirar o documento do
            # mapa, que so' depende do XML (ja guardado com sucesso acima).
            logger.warning("Falha ao descarregar PDF do IUD %s (XML OK): %s", iud, exc)

        return True

    sucesso, falhas = 0, 0
    total_a_descarregar = len(documentos_a_descarregar)
    if progresso_callback:
        progresso_callback("download", 0, max(total_a_descarregar, 1))
    with ThreadPoolExecutor(max_workers=DOWNLOAD_MAX_WORKERS) as executor:
        futuros = [executor.submit(_baixar_um, doc, direcao) for doc, direcao in documentos_a_descarregar]
        concluidos = 0
        for futuro in as_completed(futuros):
            concluidos += 1
            if futuro.result():
                sucesso += 1
            else:
                falhas += 1
            if progresso_callback:
                progresso_callback("download", concluidos, total_a_descarregar)
            if concluidos % 25 == 0 or concluidos == total_a_descarregar:
                logger.info("Progresso do download: %s/%s", concluidos, total_a_descarregar)

    logger.info("Download concluido: %s OK, %s falhas.", sucesso, falhas)
    return nome_empresa, destino_emissao, destino_rececao


def gerar_mapa(pasta: Path, mes: str, direcao: str) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Lista as faturas/faturas-recibo do mes no formato do modelo
    "Faturas controlo", sem qualquer classificacao contabilistica.

    direcao: "rececao" usa o emissor do documento (fornecedor) como
    contraparte; "emissao" usa o destinatario (cliente), ja que quem emite
    e' a propria empresa."""
    docs = carregar_documentos_de_pasta(pasta)
    campo_nome = "emissor_nome" if direcao == "rececao" else "receptor_nome"
    campo_nif = "emissor_nif" if direcao == "rececao" else "receptor_nif"
    coluna_nome = "fornecedor" if direcao == "rececao" else "cliente"
    coluna_nif = "nif_fornecedor" if direcao == "rececao" else "nif_cliente"

    linhas = []
    lancamento = 0
    total_fatura_recibo = fora_do_mapa = 0
    total_valor = 0.0
    total_base = total_iva = total_retencao = 0

    for doc in docs:
        if doc["tipo"] not in TIPOS_PROCESSAVEIS:
            fora_do_mapa += 1
            continue

        total_fatura_recibo += 1
        lancamento += 1
        prefixo = PREFIXO_TIPO.get(doc["tipo"], "FT")
        contraparte = doc[campo_nome]
        descricao_txt = f"{prefixo}{doc['numero']} {contraparte}"

        val = float(doc["valor"]) if doc["valor"] else 0.0
        base = round(float(doc["base_sem_iva"])) if doc["base_sem_iva"] else 0
        iva_v = round(float(doc["iva"])) if doc["iva"] else 0
        retencao_v = round(float(doc["retencao"])) if doc["retencao"] else 0
        total_valor += val
        total_base += base
        total_iva += iva_v
        total_retencao += retencao_v

        linhas.append({
            "lancamento_txt": lancamento,
            "tipo": prefixo,
            "data_emissao": doc["data"],
            coluna_nome: contraparte,
            coluna_nif: doc[campo_nif],
            "serie": doc["serie"],
            "numero_documento": doc["numero"],
            "descricao_txt": descricao_txt,
            "descricao_real_fatura": doc["itens"],
            "total_original": val,
            "base_sem_iva_arredondado": base,
            "iva_arredondado": iva_v,
            "retencao_arredondado": retencao_v,
            "id_efatura": doc["iud"],
        })

    df = pd.DataFrame(linhas)
    resumo = pd.DataFrame([
        (f"Documentos (FTE/FRE/TVE/NCE/NDE/DVE) {mes_nome_pt(mes)}", total_fatura_recibo),
        ("Fora do mapa (outros tipos, ex: Recibo/Transporte)", fora_do_mapa),
        ("Total arredondado", round(total_valor)),
        ("Base/valor debito arredondado", total_base),
        ("IVA manual arredondado", total_iva),
        ("Retencao total arredondada", total_retencao),
    ], columns=["Resumo", "Valor"])
    return df, resumo


def processar(
    nif: str, password: str, mes: str, incluir_emissao: bool, incluir_rececao: bool,
    logger: logging.Logger, progresso_callback=None,
) -> list[Path]:
    nome_empresa, pasta_emissao, pasta_rececao = descarregar_documentos(
        nif, password, mes, incluir_emissao, incluir_rececao, logger,
        progresso_callback=progresso_callback,
    )

    SAIDA_DIR.mkdir(parents=True, exist_ok=True)
    saidas = []

    if incluir_rececao:
        logger.info("A gerar mapa de controlo (Recebidas)...")
        df, resumo = gerar_mapa(pasta_rececao, mes, "rececao")
        out_path = SAIDA_DIR / f"Faturas controlo {sanitizar_nome(nome_empresa)} {mes} - Recebidas.xlsx"
        with pd.ExcelWriter(caminho_longo(out_path), engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="Integracao", index=False)
            resumo.to_excel(writer, sheet_name="Resumo", index=False)
        logger.info("Concluido: %s faturas recebidas -> %s", len(df), out_path)
        saidas.append(out_path)

    if incluir_emissao:
        logger.info("A gerar mapa de controlo (Emitidas)...")
        df, resumo = gerar_mapa(pasta_emissao, mes, "emissao")
        out_path = SAIDA_DIR / f"Faturas controlo {sanitizar_nome(nome_empresa)} {mes} - Emitidas.xlsx"
        with pd.ExcelWriter(caminho_longo(out_path), engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="Integracao", index=False)
            resumo.to_excel(writer, sheet_name="Resumo", index=False)
        logger.info("Concluido: %s faturas emitidas -> %s", len(df), out_path)
        saidas.append(out_path)

    return saidas


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("eFatura - Preenchimento automatico do mapa de controlo")
        self.geometry("640x480")
        self.resizable(True, True)

        self.fila_log: queue.Queue = queue.Queue()
        self.fila_progresso: queue.Queue = queue.Queue()
        self._progresso_determinado = False
        self._construir_ui()
        self.after(150, self._drenar_log)
        self.after(150, self._drenar_progresso)

    def _construir_ui(self):
        pad = {"padx": 8, "pady": 4}

        frame_form = ttk.Frame(self)
        frame_form.pack(fill="x", **pad)

        ttk.Label(frame_form, text="NIF do cliente:").grid(row=0, column=0, sticky="w")
        self.var_nif = tk.StringVar()
        ttk.Entry(frame_form, textvariable=self.var_nif, width=25).grid(row=0, column=1, sticky="w")

        ttk.Label(frame_form, text="Password do site eFatura:").grid(row=1, column=0, sticky="w")
        self.var_password = tk.StringVar()
        ttk.Entry(frame_form, textvariable=self.var_password, width=25, show="*").grid(row=1, column=1, sticky="w")

        ttk.Label(frame_form, text="Mes (AAAA-MM):").grid(row=2, column=0, sticky="w")
        self.var_mes = tk.StringVar(value=datetime.now().strftime("%Y-%m"))
        ttk.Entry(frame_form, textvariable=self.var_mes, width=25).grid(row=2, column=1, sticky="w")

        ttk.Label(frame_form, text="Gerar mapa de:").grid(row=3, column=0, sticky="w")
        frame_direcoes = ttk.Frame(frame_form)
        frame_direcoes.grid(row=3, column=1, sticky="w")
        self.var_rececao = tk.BooleanVar(value=True)
        self.var_emissao = tk.BooleanVar(value=True)
        ttk.Checkbutton(frame_direcoes, text="Receção (compras)", variable=self.var_rececao).pack(side="left")
        ttk.Checkbutton(frame_direcoes, text="Emissão (vendas)", variable=self.var_emissao).pack(side="left", padx=(10, 0))

        self.btn_processar = ttk.Button(frame_form, text="Processar", command=self._on_processar)
        self.btn_processar.grid(row=4, column=0, columnspan=2, pady=8)

        self.lbl_progresso = ttk.Label(self, text="")
        self.lbl_progresso.pack(fill="x", padx=8)

        self.progress = ttk.Progressbar(self, mode="indeterminate")
        self.progress.pack(fill="x", padx=8, pady=(0, 4))

        self.txt_log = scrolledtext.ScrolledText(self, state="disabled", wrap="word")
        self.txt_log.pack(fill="both", expand=True, padx=8, pady=4)

    def _log(self, msg: str):
        self.txt_log.configure(state="normal")
        self.txt_log.insert("end", msg + "\n")
        self.txt_log.see("end")
        self.txt_log.configure(state="disabled")

    def _drenar_log(self):
        try:
            while True:
                msg = self.fila_log.get_nowait()
                self._log(msg)
        except queue.Empty:
            pass
        self.after(150, self._drenar_log)

    def _drenar_progresso(self):
        try:
            while True:
                fase, concluidos, total = self.fila_progresso.get_nowait()
                if not self._progresso_determinado:
                    self.progress.stop()
                    self.progress.configure(mode="determinate", maximum=100)
                    self._progresso_determinado = True
                total = max(total, 1)
                pct = int(concluidos * 100 / total)
                self.progress["value"] = pct
                nome_fase = "A listar documentos" if fase == "listagem" else "A descarregar documentos"
                self.lbl_progresso.configure(text=f"{nome_fase}: {concluidos}/{total} ({pct}%)")
        except queue.Empty:
            pass
        self.after(150, self._drenar_progresso)

    def _on_processar(self):
        nif = self.var_nif.get().strip()
        password = self.var_password.get().strip()
        mes = self.var_mes.get().strip()
        incluir_rececao = self.var_rececao.get()
        incluir_emissao = self.var_emissao.get()

        if not nif or not password:
            messagebox.showwarning("Dados em falta", "Preenche o NIF e a password.")
            return
        if not re.match(r"^\d{4}-\d{2}$", mes):
            messagebox.showwarning("Mes invalido", "Usa o formato AAAA-MM, ex: 2026-05.")
            return
        if not incluir_rececao and not incluir_emissao:
            messagebox.showwarning("Nada selecionado", "Escolhe pelo menos uma direcao (Receção ou Emissão).")
            return

        self.btn_processar.configure(state="disabled")
        self._progresso_determinado = False
        self.lbl_progresso.configure(text="A autenticar no portal eFatura...")
        self.progress.configure(mode="indeterminate")
        self.progress.start(12)
        threading.Thread(
            target=self._processar_thread,
            args=(nif, password, mes, incluir_emissao, incluir_rececao),
            daemon=True,
        ).start()

    def _processar_thread(self, nif, password, mes, incluir_emissao, incluir_rececao):
        logger = criar_logger(self.fila_log)

        def progresso_cb(fase, concluidos, total):
            self.fila_progresso.put((fase, concluidos, total))

        try:
            saidas = processar(
                nif, password, mes, incluir_emissao, incluir_rececao, logger,
                progresso_callback=progresso_cb,
            )
            lista = "\n".join(str(p) for p in saidas)
            self.fila_log.put(f"\nTUDO PRONTO! Ficheiro(s) guardado(s):\n{lista}")
            self.after(0, lambda: messagebox.showinfo("Concluido", f"Mapa(s) de controlo gerado(s):\n{lista}"))
        except Exception as exc:
            self.fila_log.put(f"\nERRO: {exc}")
            self.after(0, lambda: messagebox.showerror("Erro", str(exc)))
        finally:
            self.after(0, self._reativar_botao)

    def _reativar_botao(self):
        self.progress.stop()
        self.progress.configure(mode="indeterminate")
        self._progresso_determinado = False
        self.lbl_progresso.configure(text="")
        self.btn_processar.configure(state="normal")


if __name__ == "__main__":
    App().mainloop()
