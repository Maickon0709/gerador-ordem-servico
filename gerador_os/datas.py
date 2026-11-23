"""Cálculo das datas da OS: dias úteis, feriados e prazos FOB/CIF."""

from __future__ import annotations

from datetime import date, timedelta


def pascoa(ano: int) -> date:
    """Domingo de Páscoa pelo algoritmo de Meeus/Jones/Butcher."""
    a = ano % 19
    b, c = divmod(ano, 100)
    d, e = divmod(b, 4)
    f = (b + 8) // 25
    g = (b - f + 1) // 3
    h = (19 * a + b - d - g + 15) % 30
    i, k = divmod(c, 4)
    l = (32 + 2 * e + 2 * i - h - k) % 7
    m = (a + 11 * h + 22 * l) // 451
    mes, dia = divmod(h + l - 7 * m + 114, 31)
    return date(ano, mes, dia + 1)


def feriados_fixos(ano: int) -> set[date]:
    """Feriados nacionais + data magna de Pernambuco + Sexta-feira Santa."""
    fixos = {
        (1, 1),    # Confraternização Universal
        (3, 6),    # Data Magna de Pernambuco
        (4, 21),   # Tiradentes
        (5, 1),    # Dia do Trabalho
        (9, 7),    # Independência
        (10, 12),  # Nossa Senhora Aparecida
        (11, 2),   # Finados
        (11, 15),  # Proclamação da República
        (11, 20),  # Consciência Negra
        (12, 25),  # Natal
    }
    datas = {date(ano, m, d) for m, d in fixos}
    datas.add(pascoa(ano) - timedelta(days=2))  # Sexta-feira Santa
    return datas


def somar_dias_uteis(inicio: date, dias: int, feriados: set[date]) -> date:
    """Avança `dias` dias úteis a partir de `inicio`, pulando fins de semana e feriados."""
    atual = inicio
    restantes = dias
    while restantes > 0:
        atual += timedelta(days=1)
        if atual.weekday() < 5 and atual not in feriados:
            restantes -= 1
    return atual


def calcular_datas(emissao: date, frete: str, prazo_menor: int, prazo_maior: int,
                   feriados_extras: list[date] | None = None) -> dict:
    """Regras de prazo da OS.

    - FOB: entrega = emissão + menor prazo (o cliente retira).
    - CIF: entrega = emissão + maior prazo; a OS precisa estar finalizada
      na emissão + menor prazo, para dar tempo do transporte.
    """
    feriados = feriados_fixos(emissao.year) | feriados_fixos(emissao.year + 1)
    feriados |= set(feriados_extras or [])
    menor = somar_dias_uteis(emissao, prazo_menor, feriados)
    if frete.upper() == "CIF":
        return {"entrega": somar_dias_uteis(emissao, prazo_maior, feriados), "finalizacao": menor}
    return {"entrega": menor, "finalizacao": None}


def serial_excel(d: date) -> int:
    """Converte uma data para o número de série do Excel."""
    return (d - date(1899, 12, 30)).days
