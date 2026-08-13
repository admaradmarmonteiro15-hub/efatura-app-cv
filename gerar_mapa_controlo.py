"""
Le os XML de Documentos Fiscais Eletronicos (DFE) descarregados pela
eFatura_App e devolve-os como uma lista de dicionarios prontos a usar na
geracao do mapa de controlo Excel.

Usado como biblioteca por eFatura_App.py.
"""
import glob
import html
import os
import re
from pathlib import Path

MESES_PT = {
    1: "janeiro", 2: "fevereiro", 3: "março", 4: "abril",
    5: "maio", 6: "junho", 7: "julho", 8: "agosto",
    9: "setembro", 10: "outubro", 11: "novembro", 12: "dezembro",
}


def mes_nome_pt(mes: str) -> str:
    """'2026-05' -> 'maio 2026'"""
    ano, m = mes.split("-")
    return f"{MESES_PT[int(m)]} {ano}"


def _caminho_longo(caminho: str) -> str:
    """No Windows, abrir um caminho com mais de 260 caracteres falha a
    nao ser que se use o prefixo de caminho longo \\\\?\\ - nomes de
    empresa/fornecedor compridos ultrapassam facilmente esse limite."""
    if os.name != "nt" or caminho.startswith("\\\\?\\"):
        return caminho
    return "\\\\?\\" + os.path.abspath(caminho)


def carregar_documentos_de_pasta(pasta: Path) -> list[dict]:
    docs = []
    for fp in glob.glob(str(pasta / "*.xml")):
        with open(_caminho_longo(fp), encoding="utf-8") as f:
            conteudo = f.read()

        def campo(tag, txt=conteudo):
            m = re.search(rf"<{tag}[^>]*>(.*?)</{tag}>", txt, re.S)
            return m.group(1).strip() if m else ""

        emissor_m = re.search(r"<EmitterParty>(.*?)</EmitterParty>", conteudo, re.S)
        emissor_txt = emissor_m.group(1) if emissor_m else ""

        receptor_m = re.search(r"<ReceiverParty>(.*?)</ReceiverParty>", conteudo, re.S)
        receptor_txt = receptor_m.group(1) if receptor_m else ""

        descricoes = re.findall(r"<Item>.*?<Description>(.*?)</Description>", conteudo, re.S)
        descricoes = [html.unescape(d.strip()) for d in descricoes]

        taxas = re.findall(
            r'<Tax TaxTypeCode="([^"]+)"[^>]*>\s*(?:<TaxPercentage>([\d.]+)</TaxPercentage>)?',
            conteudo,
        )
        taxa_max = max((float(p) for _, p in taxas if p), default=0.0)

        tipo = re.search(r'DocumentTypeCode="(\d+)"', conteudo)
        iud = re.search(r'<Dfe[^>]*\sId="([^"]+)"', conteudo)

        docs.append({
            "iud": iud.group(1) if iud else "",
            "tipo": tipo.group(1) if tipo else "",
            "emissor_nif": campo("TaxId", emissor_txt),
            "emissor_nome": campo("Name", emissor_txt),
            "receptor_nif": campo("TaxId", receptor_txt),
            "receptor_nome": campo("Name", receptor_txt),
            "numero": campo("DocumentNumber"),
            "serie": campo("Serie"),
            "valor": campo("PayableAmount"),
            "base_sem_iva": campo("NetTotalAmount"),
            "iva": campo("TaxTotalAmount"),
            "retencao": campo("WithholdingTaxTotalAmount"),
            "data": campo("IssueDate")[:10],
            "itens": "; ".join(descricoes),
            "isento": taxa_max == 0.0,
        })
    return docs
