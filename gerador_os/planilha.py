"""Preenchimento da planilha de Ordem de Serviço a partir do orçamento."""

from __future__ import annotations

import re
import shutil
from copy import copy
from datetime import date, datetime
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter

from .datas import calcular_datas, serial_excel
from .leitura import Orcamento

# Layout do modelo (modelo/MODELO_OS.xlsx)
PRIORIDADE_PADRAO = "II"
PRIMEIRA_LINHA_ITEM = 7
LINHAS_DE_ITEM = 6               # o modelo vem com 6 linhas; acima disso o programa insere linhas
LINHA_QUADRO = 13                # quadro grande mesclado G13:S27
LINHA_PMP = 27
TEXTO_SEGUIR_DESENHO = "*SEGUIR DESENHO*"
FAIXA_FERIADOS = "DADOS!$M$10:$M$60"

# coluna da planilha -> campo do item
COLUNAS_ITEM = {
    "B": "codigo", "C": "descricao", "G": "alt", "H": "larg", "I": "prof", "J": "uso",
    "K": "unidade", "L": "quantidade", "M": "cor", "P": "detalhes",
}


def tamanho_fonte(texto: str, caracteres_por_linha_11pt: int = 38, altura_pt: float = 34,
                  tamanho_base: int = 11) -> int:
    """Maior fonte (até o tamanho do modelo) em que o texto cabe na célula, com quebra de linha."""
    n = len(str(texto))
    for tam in range(tamanho_base, 6, -1):
        por_linha = int(caracteres_por_linha_11pt * 11 / tam)
        linhas_que_cabem = int(altura_pt / (tam * 1.3))
        if -(-n // por_linha) <= linhas_que_cabem:
            return tam
    return 7


def _numero(valor):
    """'15' ou '2,5' viram número; texto continua texto; vazio vira None."""
    texto = str(valor if valor is not None else "").strip()
    if not texto:
        return None
    if re.fullmatch(r"\d+([.,]\d+)?", texto):
        n = float(texto.replace(",", "."))
        return int(n) if n.is_integer() else n
    return texto


def montar_plano(orc: Orcamento, emissao: date, feriados_extras=None,
                 utilizar: list | None = None, pmp: str = "", seguir_desenho: bool = False) -> dict:
    """Decide o que vai em cada célula, sem tocar na planilha (fácil de testar)."""
    datas = calcular_datas(emissao, orc.frete, orc.prazo_menor, orc.prazo_maior, feriados_extras)
    extras = max(0, len(orc.itens) - LINHAS_DE_ITEM)
    valores: dict[str, object] = {
        "E1": int(orc.numero),
        "K1": emissao,
        "L2": PRIORIDADE_PADRAO,
        "B3": orc.cliente,
        "I3": orc.local,
        "L3": orc.frete,
        "Q2": datas["finalizacao"],
        # DATA ENTREGA recalcula sozinha se alguém digitar os DIAS ÚTEIS em T2
        "Q1": f'=IF(T2="",{serial_excel(datas["entrega"])},WORKDAY(K1,T2,{FAIXA_FERIADOS}))',
    }
    fontes: dict[str, int] = {}
    utilizar = utilizar or []
    for n in range(LINHAS_DE_ITEM + extras):
        linha = PRIMEIRA_LINHA_ITEM + n
        valores[f"A{linha}"] = n + 1
        item = orc.itens[n] if n < len(orc.itens) else None
        for col, campo in COLUNAS_ITEM.items():
            v = getattr(item, campo) if item else None
            if isinstance(v, float) and v.is_integer():
                v = int(v)
            valores[f"{col}{linha}"] = v if v not in ("", None) else None
            if item and col in ("C", "P") and v:
                tam = tamanho_fonte(v)
                if tam < 11:
                    fontes[f"{col}{linha}"] = tam
        valores[f"N{linha}"] = _numero(utilizar[n]) if item and n < len(utilizar) else None
    return {
        "linhas_extras": extras,
        "valores": valores,
        "fontes": fontes,
        "entrega": datas["entrega"],
        "pmp": (LINHA_PMP + extras, str(pmp or "").strip()),
        "quadro": f"G{LINHA_QUADRO + extras}" if seguir_desenho else None,
    }


def nome_arquivo(orc: Orcamento) -> str:
    cliente = re.sub(r'[\\/:*?"<>|]', "-", orc.cliente).rstrip(" .")
    return f"OS {orc.numero} - {cliente}.xlsx"


def feriados_do_modelo(modelo: str | Path) -> list[date]:
    """Datas extras da coluna FERIADOS (aba DADOS): férias coletivas, feriados municipais..."""
    wb = load_workbook(modelo, read_only=True, data_only=True)
    if "DADOS" not in wb.sheetnames:
        return []
    datas = []
    for (valor,) in wb["DADOS"].iter_rows(min_row=10, max_row=60, min_col=13, max_col=13, values_only=True):
        if isinstance(valor, datetime):
            datas.append(valor.date())
        elif isinstance(valor, date):
            datas.append(valor)
    wb.close()
    return datas


def _inserir_linhas_de_item(ws, quantidade: int) -> None:
    """Insere linhas depois da última linha de item, copiando formato e mesclagens."""
    if quantidade <= 0:
        return
    ultima = PRIMEIRA_LINHA_ITEM + LINHAS_DE_ITEM - 1
    mescladas = [m for m in ws.merged_cells.ranges]
    for m in mescladas:
        ws.unmerge_cells(str(m))
    ws.insert_rows(ultima + 1, quantidade)
    for m in mescladas:
        if m.min_row > ultima:
            m.shift(0, quantidade)
        ws.merge_cells(str(m))
    altura = ws.row_dimensions[ultima].height
    for k in range(1, quantidade + 1):
        nova = ultima + k
        ws.row_dimensions[nova].height = altura
        for col in range(1, ws.max_column + 1):
            origem = ws.cell(ultima, col)
            destino = ws.cell(nova, col)
            if origem.has_style:
                destino._style = copy(origem._style)
        for m in mescladas:
            if m.min_row == m.max_row == ultima:
                ws.merge_cells(f"{get_column_letter(m.min_col)}{nova}:{get_column_letter(m.max_col)}{nova}")


def preencher_os(modelo: str | Path, destino: str | Path, plano: dict) -> Path:
    """Copia o modelo para `destino` e escreve o plano nele."""
    destino = Path(destino)
    destino.parent.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(modelo, destino)
    wb = load_workbook(destino)
    ws = wb["MODELO"]
    _inserir_linhas_de_item(ws, plano["linhas_extras"])

    for ref, valor in plano["valores"].items():
        ws[ref].value = valor
    for ref, tam in plano["fontes"].items():
        fonte = copy(ws[ref].font)
        fonte.size = tam
        ws[ref].font = fonte

    linha_pmp, valor_pmp = plano["pmp"]
    celula = ws[f"A{linha_pmp}"]
    if isinstance(celula.value, str) and "|" in celula.value:
        celula.value = re.sub(r"\|[^|]*$", f"| {valor_pmp}", celula.value)

    if plano["quadro"]:
        q = ws[plano["quadro"]]
        q.value = TEXTO_SEGUIR_DESENHO
        q.alignment = Alignment(horizontal="center", vertical="center")
        q.font = Font(bold=True, size=28)

    ws.print_area = f"A1:S{LINHA_PMP + plano['linhas_extras']}"
    wb.save(destino)
    return destino


def gerar_os(pdf: str | Path, modelo: str | Path, pasta_saida: str | Path, emissao: date | None = None,
             utilizar=None, pmp: str = "", seguir_desenho: bool = False, substituir: bool = False) -> Path:
    """Fluxo completo: lê o PDF, monta o plano e salva a OS na pasta de saída."""
    from .leitura import ler_orcamento

    orc = ler_orcamento(pdf)
    plano = montar_plano(orc, emissao or date.today(), feriados_do_modelo(modelo),
                         utilizar=utilizar, pmp=pmp, seguir_desenho=seguir_desenho)
    saida = Path(pasta_saida) / nome_arquivo(orc)
    if saida.exists() and not substituir:
        n = 2
        while (alt := saida.with_name(f"{saida.stem} ({n}).xlsx")).exists():
            n += 1
        saida = alt
    return preencher_os(modelo, saida, plano)
