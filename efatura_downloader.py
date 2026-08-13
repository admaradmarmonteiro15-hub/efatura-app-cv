"""
Funcoes de acesso a Plataforma Eletronica (eFatura) da DNRE de Cabo Verde:
login via Playwright, listagem e download de Documentos Fiscais
Eletronicos (DFE) em XML/PDF por IUD.

Usado como biblioteca por eFatura_App.py - nao tem interface de linha de
comandos propria.
"""
import base64
import calendar
import json
import logging
import re
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime

import requests
from playwright.sync_api import sync_playwright

PE_URL = "https://pe.efatura.cv"
API_HOST = "services.efatura.cv"
DFE_LIST_URL = f"https://{API_HOST}/v2/dfe"
DFE_XML_URL = f"https://{API_HOST}/v1/dfe/xml/{{iud}}"
DFE_PDF_URL = f"https://{API_HOST}/v1/dfe/pdf/{{iud}}"

ITEMS_PER_PAGE = 100


def mes_para_intervalo(mes: str) -> tuple[str, str]:
    ano, mes_num = (int(p) for p in mes.split("-"))
    ultimo_dia = calendar.monthrange(ano, mes_num)[1]
    inicio = f"{ano:04d}-{mes_num:02d}-01"
    fim = f"{ano:04d}-{mes_num:02d}-{ultimo_dia:02d}"
    return inicio, fim


def sanitizar_nome(texto: str) -> str:
    texto = texto.strip()
    texto = re.sub(r"[^\w\s.-]", "_", texto, flags=re.UNICODE)
    texto = re.sub(r"\s+", "_", texto)
    # Nomes de pasta no Windows nao podem terminar em "." ou espaco - o Windows
    # normalmente remove isso sozinho, mas o prefixo \\?\ usado para caminhos
    # longos desativa essa limpeza automatica e a pasta fica inacessivel.
    return texto[:120].rstrip(" .") or "SemNome"


def login_e_capturar_sessao(nif: str, password: str, logger: logging.Logger) -> tuple[dict, str]:
    """Faz login real via Playwright e devolve (headers_auth, nome_empresa).

    O token Bearer nunca fica persistido em storage, so e' visto nos headers
    (em minusculas) dos pedidos que a propria app faz a services.efatura.cv.
    O nome da empresa vem do claim "name" do proprio JWT."""
    capturado = {}

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()

        def on_request(request):
            if API_HOST in request.url and "authorization" in request.headers:
                capturado["authorization"] = request.headers["authorization"]
                capturado["cv-ef-repository-code"] = request.headers.get("cv-ef-repository-code")

        page.on("request", on_request)

        page.goto(PE_URL)
        page.click("text=Login")
        page.wait_for_url("**/iam.efatura.cv/**", timeout=15000)

        page.fill("#username", nif)
        page.fill("#password", password)
        page.click("#kc-login")

        page.wait_for_url(f"{PE_URL}/**", timeout=20000)

        # Forca uma chamada a API para garantir que o token e' capturado.
        page.goto(f"{PE_URL}/dfe")
        page.wait_for_timeout(2000)

        browser.close()

    if "authorization" not in capturado:
        raise RuntimeError(
            "Nao foi possivel capturar o token de autenticacao apos o login. "
            "Verifica se o NIF/password estao corretos ou se o formulario de "
            "login mudou de estrutura."
        )

    headers_auth = {
        "authorization": capturado["authorization"],
        "cv-ef-repository-code": capturado["cv-ef-repository-code"],
        "accept": "application/json",
    }
    nome_empresa = nome_da_empresa_no_token(capturado["authorization"], nif)

    logger.info("Login OK para NIF %s (%s), token capturado.", nif, nome_empresa)
    return headers_auth, nome_empresa


def nome_da_empresa_no_token(authorization_header: str, nif: str) -> str:
    try:
        jwt = authorization_header.split(" ", 1)[1]
        payload_b64 = jwt.split(".")[1]
        payload_b64 += "=" * (-len(payload_b64) % 4)
        payload = json.loads(base64.urlsafe_b64decode(payload_b64))
        return payload.get("name") or f"NIF_{nif}"
    except Exception:
        return f"NIF_{nif}"


def get_com_retry(
    url: str, headers: dict, params: dict, logger: logging.Logger, tentativas: int = 4,
    timeout: int = 30, chave_obrigatoria: str | None = None,
) -> dict:
    """chave_obrigatoria: se indicada, uma resposta HTTP 200 cujo JSON nao
    contenha essa chave e' tratada como falha (repetivel) em vez de deixar
    rebentar mais acima com KeyError - isto acontece por vezes com respostas
    truncadas por instabilidade de rede que ainda assim conseguem ser
    interpretadas como JSON valido mas incompleto."""
    ultimo_erro = None
    for tentativa in range(1, tentativas + 1):
        try:
            resp = requests.get(url, headers=headers, params=params, timeout=timeout)
            resp.raise_for_status()
            corpo = resp.json()
            if chave_obrigatoria is not None and chave_obrigatoria not in corpo:
                raise ValueError(f"Resposta sem a chave '{chave_obrigatoria}' esperada: {corpo!r}"[:300])
            return corpo
        except (requests.exceptions.RequestException, ValueError) as exc:
            ultimo_erro = exc
            if tentativa < tentativas:
                logger.warning("Tentativa %s/%s falhou para %s (%s), a repetir...", tentativa, tentativas, url, exc)
                time.sleep(2 * tentativa)
    raise ultimo_erro


# A listagem faz agregacao no servidor sobre todos os documentos do periodo;
# em empresas com muitos documentos isso pode demorar bem mais que 30s.
LIST_TIMEOUT = 120
LIST_MAX_WORKERS = 6


def listar_documentos(
    headers_auth: dict, data_inicio: str, data_fim: str, logger: logging.Logger,
    progress_callback=None,
) -> list[dict]:
    """progress_callback(paginas_concluidas, paginas_totais), se indicado, e'
    chamado a cada pagina recebida - usado pela GUI para mostrar progresso
    real em vez de uma barra indeterminada."""
    def buscar_pagina(pagina: int) -> dict:
        params = {
            "Page": pagina,
            "ItemsPerPage": ITEMS_PER_PAGE,
            "AuthorizedDateStart": data_inicio,
            "AuthorizedDateEnd": data_fim,
            "CurrentDateTimeStamp": datetime.now().strftime("%Y%m%d%H%M%S"),
        }
        return get_com_retry(
            DFE_LIST_URL, headers_auth, params, logger, timeout=LIST_TIMEOUT, chave_obrigatoria="payload",
        )["payload"]

    payload1 = buscar_pagina(1)
    itens1 = payload1.get("items", [])
    documentos = list(itens1)
    total = payload1.get("allPagesItemCount", len(itens1))
    logger.info("Pagina 1: %s documentos (acumulado %s/%s)", len(itens1), len(documentos), total)

    if len(itens1) < ITEMS_PER_PAGE or len(documentos) >= total:
        if progress_callback:
            progress_callback(1, 1)
        return documentos

    total_paginas = -(-total // ITEMS_PER_PAGE)  # arredonda para cima
    paginas_restantes = list(range(2, total_paginas + 1))
    if progress_callback:
        progress_callback(1, total_paginas)

    resultados: dict[int, list[dict]] = {}
    with ThreadPoolExecutor(max_workers=LIST_MAX_WORKERS) as executor:
        futuros = {executor.submit(buscar_pagina, p): p for p in paginas_restantes}
        for futuro in as_completed(futuros):
            pagina = futuros[futuro]
            payload = futuro.result()
            resultados[pagina] = payload.get("items", [])
            logger.info(
                "Pagina %s: %s documentos (%s/%s paginas recebidas)",
                pagina, len(resultados[pagina]), len(resultados), len(paginas_restantes),
            )
            if progress_callback:
                progress_callback(1 + len(resultados), total_paginas)

    for pagina in paginas_restantes:
        documentos.extend(resultados.get(pagina, []))

    return documentos


def descarregar_xml(headers_auth: dict, iud: str, logger: logging.Logger) -> bytes:
    body = get_com_retry(DFE_XML_URL.format(iud=iud), headers_auth, {}, logger, chave_obrigatoria="payload")
    return body["payload"].encode("utf-8")


def descarregar_pdf(headers_auth: dict, iud: str, logger: logging.Logger) -> bytes:
    """O endpoint devolve JSON {"payload": "<PDF em base64>"} - descoberto
    por interceção de rede ao clicar o botao de PDF no portal (o mesmo
    bearer token usado para XML funciona aqui sem sessao de browser)."""
    body = get_com_retry(DFE_PDF_URL.format(iud=iud), headers_auth, {}, logger, chave_obrigatoria="payload")
    return base64.b64decode(body["payload"])
