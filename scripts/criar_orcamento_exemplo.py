"""Gera orçamentos em PDF FICTÍCIOS no mesmo layout de texto do relatório do ERP.

    python scripts/criar_orcamento_exemplo.py
"""

from pathlib import Path

from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas

PASTA = Path(__file__).resolve().parent.parent / "exemplos"


def escrever(nome: str, linhas: list[str]) -> None:
    PASTA.mkdir(exist_ok=True)
    c = canvas.Canvas(str(PASTA / nome), pagesize=A4)
    largura, altura = A4
    y = altura - 50
    for linha in linhas:
        if y < 60:
            c.showPage()
            y = altura - 50
        negrito = linha.isupper() and len(linha) < 40
        c.setFont("Helvetica-Bold" if negrito else "Helvetica", 9)
        c.drawString(40, y, linha)
        y -= 14
    c.save()


CABECALHO = [
    "BETA INDUSTRIA METALURGICA LTDA",
    "CNPJ 00.000.000/0001-00 Rua: Rua Exemplo, 100 - Distrito Industrial",
    "Orçamento 004827 Data 12/11/2026",
    "Data: 12/11/2026",
    "Vendedor 004 ANA SOUZA",
    "Cliente 1203 CONSTRUTORA HORIZONTE LTDA Fantasia HORIZONTE",
    "Cidade Recife UF PE Fone (81) 0000-0000",
    "It. Código Descrição NCM Qtde Un. Vl. Unit. Total",
]

RODAPE = [
    "Observações",
    "PRAZO DE ENTREGA: 3/5 DIAS ÚTEIS",
    "FRETE: CIF",
    "Total R$ 18.450,00",
]


def orcamento_pequeno() -> None:
    itens = [
        "1 QDCOM01 QUADRO DE COMANDO 600X400X200 MM 85371090 4,00 CX R$ 1.250,00 R$ 5.000,00",
        "CINZA RAL 7035 - DE SOBREPOR - FECHO COM CHAVE",
        "2 TPSHF01 TAMPA SHAFT 1000X500X60 MM 73089010 2,00 PÇ R$ 640,00 R$ 1.280,00",
        "CINZA RAL 7035",
        "3 CXPAS01 CAIXA DE PASSAGEM 85389090 10,00 CX R$ 230,00 R$ 2.300,00",
        "300X300X120 - GALVANIZADO - DE EMBUTIR",
        "4 PNCOM01 PAINEL DE COMANDO COM PORTA INTERNA E PLACA DE MONTAGEM",
        "85371090 1,00 CJ 3.870,00 R$ 3.870,00",
        "800X600X250 - BRANCO - VISOR EM POLICARBONATO - DISPOSITIVO PARA CADEADO",
    ]
    escrever("orcamento_exemplo.pdf", CABECALHO + itens + RODAPE)


def orcamento_grande() -> None:
    itens = []
    for n in range(1, 9):
        itens.append(f"{n} CXMED0{n} CAIXA DE MEDIÇÃO TIPO {n} {300 + 50 * n}X{250 + 20 * n}X120 MM "
                     f"85389090 {n},00 CX R$ 410,00 R$ {410 * n},00")
        itens.append("CINZA RAL 7032 - DE SOBREPOR - PADRÃO CONCESSIONÁRIA")
    cab = [l.replace("004827", "004830").replace("CONSTRUTORA HORIZONTE LTDA", "ENGENHARIA EPSILON LTDA")
           for l in CABECALHO]
    rod = [l.replace("FRETE: CIF", "FRETE: FOB").replace("3/5", "10/15") for l in RODAPE]
    escrever("orcamento_exemplo_8_itens.pdf", cab + itens + rod)


if __name__ == "__main__":
    orcamento_pequeno()
    orcamento_grande()
    print(f"PDFs criados em {PASTA}")
