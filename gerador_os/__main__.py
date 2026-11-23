"""Uso:
    python -m gerador_os                      abre a janela
    python -m gerador_os orcamento.pdf        gera direto pela linha de comando
"""

import argparse
from pathlib import Path

from .leitura import OrcamentoInvalido
from .planilha import gerar_os

RAIZ = Path(__file__).resolve().parent.parent
MODELO_PADRAO = RAIZ / "modelo" / "MODELO_OS.xlsx"


def main() -> int:
    p = argparse.ArgumentParser(prog="gerador_os", description="Gera a Ordem de Serviço a partir do orçamento em PDF.")
    p.add_argument("pdf", nargs="?", help="orçamento em PDF (sem ele, abre a janela)")
    p.add_argument("--modelo", default=str(MODELO_PADRAO), help="planilha modelo da OS")
    p.add_argument("--saida", default="OS geradas", help="pasta onde a OS é salva")
    p.add_argument("--utilizar", nargs="*", default=[], help="estoque a utilizar, um valor por item")
    p.add_argument("--pmp", default="", help="valor do PMP")
    p.add_argument("--seguir-desenho", action="store_true", help="escreve *SEGUIR DESENHO* no quadro grande")
    p.add_argument("--substituir", action="store_true", help="substitui a OS se já existir uma com o mesmo número")
    args = p.parse_args()

    if not args.pdf:
        from .janela import abrir_janela
        abrir_janela(Path(args.modelo))
        return 0

    try:
        saida = gerar_os(args.pdf, args.modelo, args.saida, utilizar=args.utilizar, pmp=args.pmp,
                         seguir_desenho=args.seguir_desenho, substituir=args.substituir)
    except OrcamentoInvalido as erro:
        print(f"Erro: {erro}")
        return 1
    print(f"OS gerada: {saida}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
