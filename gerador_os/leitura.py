"""Leitura do orçamento em PDF exportado pelo ERP.

O texto do PDF é reconstruído linha a linha (agrupando palavras pela altura na
página) e depois interpretado com expressões regulares: cabeçalho (número,
cliente, cidade, frete, prazo) e itens (código, descrição, medidas,
quantidade, unidade, cor, uso e detalhes).
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path


class OrcamentoInvalido(ValueError):
    """O PDF não tem alguma informação obrigatória do orçamento."""


@dataclass
class Item:
    numero: int
    codigo: str
    descricao: str
    alt: float | None
    larg: float | None
    prof: float | None
    unidade: str
    quantidade: float
    uso: str = ""
    cor: str = ""
    detalhes: str = ""


@dataclass
class Orcamento:
    numero: str
    data_emissao: str | None
    cliente: str
    local: str
    frete: str
    prazo_menor: int
    prazo_maior: int
    itens: list[Item] = field(default_factory=list)


# ---------------------------------------------------------------- texto do PDF

def texto_do_pdf(caminho: str | Path) -> str:
    """Extrai o texto do PDF juntando as palavras que estão na mesma linha visual."""
    import pdfplumber

    linhas: list[str] = []
    with pdfplumber.open(caminho) as pdf:
        for pagina in pdf.pages:
            palavras = sorted(pagina.extract_words(keep_blank_chars=False),
                              key=lambda w: (round(w["top"], 1), w["x0"]))
            grupos: list[list[dict]] = []
            for w in palavras:
                if grupos and abs(grupos[-1][0]["top"] - w["top"]) < 2.5:
                    grupos[-1].append(w)
                else:
                    grupos.append([w])
            for g in grupos:
                g.sort(key=lambda w: w["x0"])
                linhas.append(re.sub(r"\s+", " ", " ".join(w["text"] for w in g)).strip())
    return "\n".join(linhas)


# ---------------------------------------------------------- padrões do relatório

RE_DATA = re.compile(r"^Data:\s*(\d{2}/\d{2}/\d{4})", re.M)
RE_NUMERO = re.compile(r"Or[çc]amento\s+0*(\d+)\s+Data", re.I)
RE_CLIENTE = re.compile(r"^Cliente\s+\d+\s+(.+?)\s+Fantasia", re.I | re.M)
RE_CIDADE = re.compile(r"^Cidade\s+(.+?)\s+UF\s+([A-Z]{2})\b", re.I | re.M)
RE_PRAZO = re.compile(r"PRAZO\s+DE\s+ENTREGA\s*:?\s*(\d+)\s*/\s*(\d+)", re.I)
RE_FRETE = re.compile(r"FRETE\s*:?\s*(CIF|FOB)", re.I)

# Item em uma linha:  1  QDCOM  QUADRO DE COMANDO 600X400X200  73089010  4,00  CX  R$ ...
RE_ITEM = re.compile(r"^(\d+)\s+(\S+)\s+(.+?)\s+(\d{8})\s+([\d.,]+)\s+(\S+)\s+R\$")
# Item quebrado em duas linhas (descrição longa): a 2ª linha começa no NCM
RE_ITEM_INICIO = re.compile(r"^(\d+)\s+(\S+)\s+(.+)$")
RE_ITEM_RESTO = re.compile(r"^(\d{8})\s+(?:\d{7}\s+)?([\d.,]+)\s+(\S+)\s+[\d.,]+\s+R\$")

NUM = r"(\d+(?:[.,]\d+)?)"
RE_MEDIDAS_FIM = re.compile(rf"\s*{NUM}\s*X\s*{NUM}(?:\s*X\s*{NUM})?\s*(?:MM)?\s*$", re.I)
RE_MEDIDAS = re.compile(rf"^{NUM}\s*X\s*{NUM}(?:\s*X\s*{NUM})?\s*(?:MM)?$", re.I)

RE_FIM_ITENS = re.compile(r"^(Observa[çc][õo]es|Total\b|Endere[çc]o (Entrega|Cobran))")
RE_CABECALHO = re.compile(
    r"^(BETA INDUSTRIA|CNPJ|Rua:|Complemento:|UF:|Email|Data:|Or[çc]amento|Vendedor|Cliente|"
    r"Endere[çc]o|Cidade|Fone|Contato|Transportadora|It\.|P[áa]gina\b)"
)
RE_COR = re.compile(r"\b(CINZA|BRANC[OA]|PRET[OA]|AZUL|VERDE|VERMELH[OA]|AMAREL[OA]|LARANJA|BEGE|"
                    r"MARROM|GRAFITE|PRATA|RAL\s*\d+|MUNSELL)\b")
RE_ACABAMENTO = re.compile(r"\b(GALVANIZAD[OA]|NATURAL|INOX)\b")
RE_USO = re.compile(r"(?:^|\bDE\s+)(EMBUTI(?:R|DO|DA)|SOBREP(?:OR|OSTO|OSTA))\b", re.I)


def numero_br(texto: str) -> float:
    """'1.234,50' -> 1234.5"""
    return float(texto.replace(".", "").replace(",", "."))


def _uso(palavra: str) -> str:
    return "EMBUTIR" if palavra.upper().startswith("EMB") else "SOBREPOR"


def _medidas(grupos: tuple) -> tuple:
    return tuple(numero_br(g) if g else None for g in grupos)


def separar_dimensoes(descricao: str) -> dict:
    """'QUADRO 600X400X200 MM' -> descrição 'QUADRO' + alt/larg/prof."""
    m = RE_MEDIDAS_FIM.search(descricao)
    if not m:
        return {"descricao": descricao.strip(), "alt": None, "larg": None, "prof": None}
    alt, larg, prof = _medidas(m.groups())
    return {"descricao": descricao[: m.start()].strip(), "alt": alt, "larg": larg, "prof": prof}


def _medidas_no_complemento(linhas: list[str]):
    """Procura medidas nas linhas de complemento ('CINZA - 600X400 - CH#20')."""
    for i, linha in enumerate(linhas):
        partes = re.split(r"\s+-\s+", linha)
        for j, parte in enumerate(partes):
            m = RE_MEDIDAS.match(parte.strip())
            if m:
                del partes[j]
                novas = list(linhas)
                if partes:
                    novas[i] = " - ".join(partes)
                else:
                    del novas[i]
                alt, larg, prof = _medidas(m.groups())
                return {"alt": alt, "larg": larg, "prof": prof}, novas
    return None, linhas


def separar_cor_detalhes(linhas: list[str]) -> dict:
    """Do complemento do item separa a cor, o uso (EMBUTIR/SOBREPOR) e o resto (detalhes)."""
    partes = [p.strip() for linha in linhas for p in re.split(r"\s+-\s+", linha) if p.strip()]
    uso = ""
    for i, parte in enumerate(partes):
        m = RE_USO.search(parte)
        if m:
            uso = _uso(m.group(1))
            partes[i] = re.sub(r"\s+", " ", parte[: m.start()] + " " + parte[m.end():]).strip()
            break
    partes = [p for p in partes if p]
    idx = next((i for i, p in enumerate(partes) if RE_COR.search(p)), -1)
    if idx == -1:
        idx = next((i for i, p in enumerate(partes) if RE_ACABAMENTO.search(p)), -1)
    cor = partes[idx] if idx != -1 else ""
    detalhes = " - ".join(p for i, p in enumerate(partes) if i != idx)
    return {"cor": cor, "uso": uso, "detalhes": detalhes}


def _inicio_item(linhas: list[str], i: int):
    m = RE_ITEM.match(linhas[i])
    if m:
        return {"numero": m[1], "codigo": m[2], "descricao": m[3], "quantidade": m[5],
                "unidade": m[6], "proxima": i + 1}
    m1 = RE_ITEM_INICIO.match(linhas[i])
    m2 = m1 and i + 1 < len(linhas) and RE_ITEM_RESTO.match(linhas[i + 1])
    if m2:
        return {"numero": m1[1], "codigo": m1[2], "descricao": m1[3], "quantidade": m2[2],
                "unidade": m2[3], "proxima": i + 2}
    return None


def interpretar_texto(texto: str) -> Orcamento:
    """Transforma o texto do PDF em um Orcamento. Lança OrcamentoInvalido se faltar algo."""
    faltando: list[str] = []

    def achar(regex, nome):
        m = regex.search(texto)
        if not m:
            faltando.append(nome)
        return m

    numero = achar(RE_NUMERO, "número do orçamento")
    cliente = achar(RE_CLIENTE, "cliente")
    cidade = achar(RE_CIDADE, "cidade/UF")
    prazo = achar(RE_PRAZO, "prazo de entrega (ex.: 3/5 DIAS ÚTEIS)")
    frete = achar(RE_FRETE, "frete (CIF/FOB)")

    linhas = [l.strip() for l in texto.split("\n") if l.strip()]
    itens: list[Item] = []
    i = 0
    while i < len(linhas):
        inicio = _inicio_item(linhas, i)
        if not inicio:
            i += 1
            continue
        dims = separar_dimensoes(inicio["descricao"])
        complemento = []
        for j in range(inicio["proxima"], len(linhas)):
            if _inicio_item(linhas, j) or RE_FIM_ITENS.match(linhas[j]):
                break
            if not RE_CABECALHO.match(linhas[j]):
                complemento.append(linhas[j])
        if dims["alt"] is None:
            achadas, complemento = _medidas_no_complemento(complemento)
            if achadas:
                dims.update(achadas)
        extra = separar_cor_detalhes(complemento)
        uso = extra["uso"]
        if not uso:
            m = RE_USO.search(dims["descricao"])
            uso = _uso(m.group(1)) if m else ""
        itens.append(Item(
            numero=int(inicio["numero"]), codigo=inicio["codigo"], descricao=dims["descricao"],
            alt=dims["alt"], larg=dims["larg"], prof=dims["prof"], unidade=inicio["unidade"],
            quantidade=numero_br(inicio["quantidade"]), uso=uso, cor=extra["cor"],
            detalhes=extra["detalhes"],
        ))
        i = inicio["proxima"]

    if not itens:
        faltando.append("itens")
    if faltando:
        raise OrcamentoInvalido("Não consegui encontrar no orçamento: " + ", ".join(faltando))

    menor, maior = sorted((int(prazo[1]), int(prazo[2])))
    data = RE_DATA.search(texto)
    return Orcamento(
        numero=numero[1],
        data_emissao=data[1] if data else None,
        cliente=cliente[1].strip(),
        local=f"{cidade[1].strip()}/{cidade[2]}".upper(),
        frete=frete[1].upper(),
        prazo_menor=menor,
        prazo_maior=maior,
        itens=itens,
    )


def ler_orcamento(caminho: str | Path) -> Orcamento:
    return interpretar_texto(texto_do_pdf(caminho))
