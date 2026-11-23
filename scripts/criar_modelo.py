"""Gera o modelo genérico de OS (modelo/MODELO_OS.xlsx).

Mesma estrutura de células usada pelo gerador. Rode de novo se quiser recriar o modelo:
    python scripts/criar_modelo.py
"""

from datetime import date
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

AZUL = PatternFill("solid", fgColor="1F3B57")
CINZA = PatternFill("solid", fgColor="E7ECF1")
AMARELO = PatternFill("solid", fgColor="FFF2B3")
FINO = Side(style="thin", color="8A96A3")
BORDA = Border(left=FINO, right=FINO, top=FINO, bottom=FINO)
CENTRO = Alignment(horizontal="center", vertical="center", wrap_text=True)
ESQ = Alignment(horizontal="left", vertical="center", wrap_text=True)


def rotulo(ws, ref, texto, mescla=None, fill=CINZA, cor="000000"):
    ws[ref] = texto
    ws[ref].font = Font(bold=True, size=10, color=cor)
    ws[ref].fill = fill
    ws[ref].alignment = CENTRO
    if mescla:
        ws.merge_cells(mescla)


def campo(ws, faixa, fmt=None, fill=None):
    ws.merge_cells(faixa) if ":" in faixa else None
    ref = faixa.split(":")[0]
    ws[ref].alignment = CENTRO
    ws[ref].font = Font(size=11)
    if fmt:
        ws[ref].number_format = fmt
    if fill:
        ws[ref].fill = fill


def criar(caminho: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "MODELO"
    larguras = {"A": 9, "B": 11, "C": 10, "D": 10, "E": 10, "F": 10, "G": 9, "H": 9, "I": 9, "J": 11,
                "K": 13, "L": 8, "M": 14, "N": 9, "O": 7, "P": 10, "Q": 10, "R": 10, "S": 10, "T": 8, "U": 8}
    for col, w in larguras.items():
        ws.column_dimensions[col].width = w

    # Cabeçalho
    rotulo(ws, "A1", "EMPRESA\nFICTÍCIA", "A1:B2", fill=AZUL, cor="FFFFFF")
    rotulo(ws, "C1", "ORDEM DE SERVIÇO:", "C1:D1")
    campo(ws, "E1:G1")
    ws["E1"].font = Font(bold=True, size=14)
    rotulo(ws, "H1", "DATA EMISSÃO:", "H1:J1")
    campo(ws, "K1:M1", "dd/mm/yyyy")
    ws["K1"] = "=TODAY()"
    rotulo(ws, "N1", "DATA ENTREGA:", "N1:P1")
    campo(ws, "Q1:S1", "dd/mm/yyyy")
    ws["Q1"] = "=WORKDAY(K1,T2,DADOS!$M$10:$M$60)"
    rotulo(ws, "T1", "DIAS ÚTEIS", "T1:U1")
    campo(ws, "T2:U2", fill=AMARELO)

    rotulo(ws, "C2", "PEDIDO DE VENDA:", "C2:D2")
    campo(ws, "E2:G2")
    rotulo(ws, "H2", "MARGEM:")
    campo(ws, "I2:J2")
    rotulo(ws, "K2", "PRIORIDADE:")
    campo(ws, "L2:M2")
    rotulo(ws, "N2", "DATA FINALIZ. OS:", "N2:P2")
    campo(ws, "Q2:S2", "dd/mm/yyyy")

    rotulo(ws, "A3", "CLIENTE:")
    campo(ws, "B3:D3")
    ws["B3"].alignment = ESQ
    rotulo(ws, "E3", "OBRA:")
    campo(ws, "F3:G3")
    rotulo(ws, "H3", "LOCAL:")
    campo(ws, "I3:J3")
    rotulo(ws, "K3", "FRETE:")
    campo(ws, "L3:M3")
    rotulo(ws, "N3", "VOLUME:")
    campo(ws, "O3:P3")
    rotulo(ws, "Q3", "PESO:")
    campo(ws, "R3:S3")
    ws.merge_cells("A4:S4")

    # Tabela de itens
    for ref, texto, mescla in [("A5", "ITEM", "A5:A6"), ("B5", "CÓD.", "B5:B6"),
                               ("C5", "DESCRIÇÃO DO PRODUTO", "C5:F6"), ("G5", "DIMENSÕES (mm)", "G5:I5"),
                               ("J5", "USO", "J5:J6"), ("K5", "UNID", "K5:K6"), ("L5", "QTDE", "L5:L6"),
                               ("M5", "COR", "M5:M6"), ("N5", "ESTOQUE", "N5:O5"), ("P5", "DETALHES", "P5:S6")]:
        rotulo(ws, ref, texto, mescla, fill=AZUL, cor="FFFFFF")
    for ref, texto in [("G6", "ALT"), ("H6", "LARG"), ("I6", "PROF"), ("N6", "UTILIZ."), ("O6", "GER.")]:
        rotulo(ws, ref, texto, fill=AZUL, cor="FFFFFF")
    for linha in range(7, 13):
        ws.row_dimensions[linha].height = 34
        ws.merge_cells(f"C{linha}:F{linha}")
        ws.merge_cells(f"P{linha}:S{linha}")
        for col in "ABCDEFGHIJKLMNOPQRS":
            c = ws[f"{col}{linha}"]
            c.border = BORDA
            c.alignment = ESQ if col in "CP" else CENTRO
            c.font = Font(size=11)
        ws[f"N{linha}"].fill = AMARELO
        ws[f"A{linha}"] = linha - 6

    # Acompanhamento da produção + quadro grande
    etapas = [("DIRETORIA", None), ("AUTORIZ. PROD.", None), ("PROJETO", "CORTE"),
              ("RECEBIMENTO", "RECEBIMENTO"), ("FINALIZAÇÃO", "FINALIZAÇÃO"), ("DOBRA", "SOLDA"),
              ("RECEBIMENTO", "RECEBIMENTO"), ("FINALIZAÇÃO", "FINALIZAÇÃO"), ("PINTURA", "MONTAGEM"),
              ("RECEBIMENTO", "RECEBIMENTO"), ("FINALIZAÇÃO", "FINALIZAÇÃO"), ("COMERCIAL", None),
              ("RECEBIMENTO", None)]
    for k, (esq, dir_) in enumerate(etapas):
        linha = 13 + k
        if dir_:
            rotulo(ws, f"A{linha}", esq, f"A{linha}:C{linha}")
            rotulo(ws, f"D{linha}", dir_, f"D{linha}:F{linha}")
        else:
            rotulo(ws, f"A{linha}", esq, f"A{linha}:F{linha}")
    rotulo(ws, "A26", "PROJ | CORT | DOB | SOLD | ACAB | PINT | MONT | PMP:", "A26:F26")
    ws["A27"] = "|      |      |      |      |      |      |      | "
    ws.merge_cells("A27:F27")
    ws["A27"].alignment = CENTRO
    ws.merge_cells("G13:S27")
    ws["G13"].border = BORDA

    # Aba de apoio: feriados extras (férias coletivas, feriados municipais...)
    dados = wb.create_sheet("DADOS")
    dados.column_dimensions["L"].width = 28
    dados.column_dimensions["M"].width = 14
    dados["L8"] = "FERIADOS EXTRAS (além dos nacionais e de PE, que o programa já conhece)"
    dados.merge_cells("L8:M8")
    dados["L9"], dados["M9"] = "DESCRIÇÃO", "DATA"
    exemplos = [("FÉRIAS COLETIVAS", date(2026, 12, 28)), ("FÉRIAS COLETIVAS", date(2026, 12, 29)),
                ("FÉRIAS COLETIVAS", date(2026, 12, 30)), ("CARNAVAL", date(2027, 2, 8)),
                ("CARNAVAL", date(2027, 2, 9)), ("FERIADO MUNICIPAL (exemplo)", date(2027, 7, 16))]
    for k, (nome, d) in enumerate(exemplos):
        dados[f"L{10 + k}"] = nome
        dados[f"M{10 + k}"] = d
        dados[f"M{10 + k}"].number_format = "dd/mm/yyyy"

    ws.print_area = "A1:S27"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    caminho.parent.mkdir(parents=True, exist_ok=True)
    wb.save(caminho)


if __name__ == "__main__":
    destino = Path(__file__).resolve().parent.parent / "modelo" / "MODELO_OS.xlsx"
    criar(destino)
    print(f"Modelo criado em {destino}")
