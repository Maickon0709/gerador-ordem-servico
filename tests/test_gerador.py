from datetime import date
from pathlib import Path

import pytest
from openpyxl import load_workbook

from gerador_os.datas import calcular_datas, feriados_fixos, pascoa, somar_dias_uteis
from gerador_os.leitura import (OrcamentoInvalido, interpretar_texto, ler_orcamento, separar_cor_detalhes,
                                separar_dimensoes)
from gerador_os.planilha import feriados_do_modelo, gerar_os, montar_plano, nome_arquivo, tamanho_fonte

RAIZ = Path(__file__).resolve().parent.parent
MODELO = RAIZ / "modelo" / "MODELO_OS.xlsx"
PDF = RAIZ / "exemplos" / "orcamento_exemplo.pdf"
PDF_8 = RAIZ / "exemplos" / "orcamento_exemplo_8_itens.pdf"

TEXTO = """Orçamento 000123 Data 05/11/2026
Data: 05/11/2026
Cliente 77 OFICINA DO ZE LTDA Fantasia ZE
Cidade Jaboatao dos Guararapes UF PE Fone 0
1 TAMPA01 TAMPA DE INSPECAO 500X500 MM 73089010 3,00 PÇ R$ 100,00 R$ 300,00
PRETO - DE EMBUTIR
Observações
PRAZO DE ENTREGA: 7/2 DIAS UTEIS
FRETE: FOB"""


# ------------------------------------------------------------------ datas

def test_pascoa_conhecida():
    assert pascoa(2026) == date(2026, 4, 5)
    assert pascoa(2027) == date(2027, 3, 28)


def test_feriados_incluem_sexta_santa_e_data_magna_pe():
    f = feriados_fixos(2026)
    assert date(2026, 4, 3) in f      # Sexta-feira Santa
    assert date(2026, 3, 6) in f      # Data Magna de Pernambuco
    assert date(2026, 11, 20) in f    # Consciência Negra


def test_dias_uteis_pulam_fim_de_semana_e_feriado():
    # quinta 19/11/2026 + 1 dia útil: sexta 20/11 é feriado, então vai para segunda 23/11
    assert somar_dias_uteis(date(2026, 11, 19), 1, feriados_fixos(2026)) == date(2026, 11, 23)


def test_fob_entrega_no_menor_prazo():
    d = calcular_datas(date(2026, 11, 12), "FOB", 3, 5)
    assert d == {"entrega": date(2026, 11, 17), "finalizacao": None}


def test_cif_entrega_no_maior_prazo_e_finaliza_no_menor():
    d = calcular_datas(date(2026, 11, 12), "CIF", 3, 5)
    assert d == {"entrega": date(2026, 11, 19), "finalizacao": date(2026, 11, 17)}


def test_feriado_extra_da_empresa_empurra_a_data():
    sem = calcular_datas(date(2026, 12, 23), "FOB", 2, 2)["entrega"]
    com = calcular_datas(date(2026, 12, 23), "FOB", 2, 2, [date(2026, 12, 28)])["entrega"]
    assert sem == date(2026, 12, 28) and com == date(2026, 12, 29)


# ---------------------------------------------------------------- leitura

def test_separar_dimensoes():
    assert separar_dimensoes("QUADRO 600X400X200 MM") == {
        "descricao": "QUADRO", "alt": 600, "larg": 400, "prof": 200}
    assert separar_dimensoes("TAMPA SEM MEDIDA")["alt"] is None


def test_separar_cor_uso_detalhes():
    r = separar_cor_detalhes(["CINZA RAL 7035 - DE SOBREPOR - FECHO COM CHAVE"])
    assert r == {"cor": "CINZA RAL 7035", "uso": "SOBREPOR", "detalhes": "FECHO COM CHAVE"}


def test_interpretar_texto_cabecalho_e_item():
    o = interpretar_texto(TEXTO)
    assert (o.numero, o.cliente, o.local, o.frete) == ("123", "OFICINA DO ZE LTDA",
                                                     "JABOATAO DOS GUARARAPES/PE", "FOB")
    assert (o.prazo_menor, o.prazo_maior) == (2, 7)   # ordem do prazo no PDF não importa
    it = o.itens[0]
    assert (it.codigo, it.descricao, it.alt, it.larg, it.prof) == ("TAMPA01", "TAMPA DE INSPECAO", 500, 500, None)
    assert (it.quantidade, it.unidade, it.cor, it.uso) == (3, "PÇ", "PRETO", "EMBUTIR")


def test_orcamento_incompleto_avisa_o_que_falta():
    with pytest.raises(OrcamentoInvalido, match="frete"):
        interpretar_texto(TEXTO.replace("FRETE: FOB", ""))


def test_ler_pdf_de_exemplo():
    o = ler_orcamento(PDF)
    assert o.numero == "4827" and o.frete == "CIF" and len(o.itens) == 4
    # item 3 tem as medidas no complemento; item 4 tem a descrição quebrada em duas linhas
    assert (o.itens[2].alt, o.itens[2].uso, o.itens[2].cor) == (300, "EMBUTIR", "GALVANIZADO")
    assert o.itens[3].descricao.startswith("PAINEL DE COMANDO") and o.itens[3].quantidade == 1


# ---------------------------------------------------------------- planilha

def test_plano_preenche_celulas_e_limpa_linhas_vazias():
    o = interpretar_texto(TEXTO)
    p = montar_plano(o, date(2026, 11, 12), utilizar=["2"], pmp="3")
    v = p["valores"]
    assert v["E1"] == 123 and v["L3"] == "FOB" and v["Q2"] is None
    assert v["B7"] == "TAMPA01" and v["N7"] == 2
    assert v["B8"] is None and v["A12"] == 6
    assert p["linhas_extras"] == 0 and p["pmp"] == (27, "3")


def test_nome_do_arquivo_sem_caracteres_invalidos():
    o = interpretar_texto(TEXTO.replace("OFICINA DO ZE LTDA", "A/B: C*D LTDA."))
    assert nome_arquivo(o) == "OS 123 - A-B- C-D LTDA.xlsx"


def test_fonte_diminui_para_texto_longo():
    assert tamanho_fonte("CURTO") == 11
    assert tamanho_fonte("X" * 200) < 11


def test_feriados_extras_lidos_do_modelo():
    assert date(2026, 12, 28) in feriados_do_modelo(MODELO)


def test_gerar_os_completa(tmp_path):
    saida = gerar_os(PDF, MODELO, tmp_path, emissao=date(2026, 11, 12), utilizar=["2"], pmp="1,5",
                     seguir_desenho=True)
    ws = load_workbook(saida)["MODELO"]
    assert ws["E1"].value == 4827 and ws["B3"].value == "CONSTRUTORA HORIZONTE LTDA"
    assert ws["C7"].value == "QUADRO DE COMANDO" and ws["G7"].value == 600 and ws["J7"].value == "SOBREPOR"
    assert ws["Q1"].value.startswith('=IF(T2=""')
    assert ws["A27"].value.endswith("| 1,5")
    assert ws["G13"].value == "*SEGUIR DESENHO*"


def test_mais_de_seis_itens_insere_linhas(tmp_path):
    saida = gerar_os(PDF_8, MODELO, tmp_path, emissao=date(2026, 11, 12))
    ws = load_workbook(saida)["MODELO"]
    assert [ws[f"A{l}"].value for l in range(7, 15)] == list(range(1, 9))
    assert ws["B14"].value == "CXMED08"
    assert ws["A15"].value == "DIRETORIA"          # bloco de baixo desceu 2 linhas
    assert "G15:S29" in {str(m) for m in ws.merged_cells.ranges}


def test_os_repetida_nao_sobrescreve(tmp_path):
    a = gerar_os(PDF, MODELO, tmp_path, emissao=date(2026, 11, 12))
    b = gerar_os(PDF, MODELO, tmp_path, emissao=date(2026, 11, 12))
    assert a != b and b.name.endswith("(2).xlsx")
