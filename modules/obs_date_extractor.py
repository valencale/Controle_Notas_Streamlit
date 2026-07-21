"""
obs_date_extractor.py — Extrai datas de faturamento e entrega do campo OBS.

Padrões reconhecidos:
    Fatura: FAT 30/7, FATURA 20/07, FATURAR 15/08/2026, FAT. 10/5
    Entrega: ENTREGA EMBU 06/8, ENTREGA 23/07/2026, ENT 14/5, ENT. 01/06

Retorna datetime.date nativos para compatibilidade com openpyxl/Excel.
"""

import re
from datetime import date, datetime


def extract_dates_from_obs(obs_text: str) -> dict:
    """
    Extrai DATA_FATURA e DATA_ENTREGA do texto de observações.

    Args:
        obs_text: Conteúdo do campo OBS (ex: "FAT 30/7 ENTREGA EMBU 06/8")

    Returns:
        Dict com chaves DATA_FATURA e DATA_ENTREGA, cada uma datetime.date ou None.
    """
    if not isinstance(obs_text, str) or not obs_text.strip():
        return {"DATA FATURA": None, "DATA ENTREGA": None}

    resultado = {"DATA FATURA": None, "DATA ENTREGA": None}

    # Padrões regex: ignora palavras intermediárias (como EMBU, URGENTE)
    # até encontrar a data no formato dd/mm ou dd/mm/aaaa
    padrao_fatura = r"(?i)FAT(?:URA|URAR|\.)?\s*.*?(\d{1,2}[/-]\d{1,2}(?:[/-]\d{2,4})?)"
    padrao_entrega = r"(?i)ENT(?:REGA|\.)?\s*.*?(\d{1,2}[/-]\d{1,2}(?:[/-]\d{2,4})?)"

    # Busca Fatura
    match_fatura = re.search(padrao_fatura, obs_text)
    if match_fatura:
        resultado["DATA FATURA"] = _parse_to_date(match_fatura.group(1))

    # Busca Entrega
    match_entrega = re.search(padrao_entrega, obs_text)
    if match_entrega:
        resultado["DATA ENTREGA"] = _parse_to_date(match_entrega.group(1))

    # Regra de virada de ano: se entrega < fatura e mês entrega < mês fatura,
    # provavelmente a entrega é no ano seguinte
    if resultado["DATA FATURA"] and resultado["DATA ENTREGA"]:
        fat = resultado["DATA FATURA"]
        ent = resultado["DATA ENTREGA"]
        if ent.month < fat.month and fat.month >= 11:
            resultado["DATA ENTREGA"] = date(ent.year + 1, ent.month, ent.day)

    return resultado


def _parse_to_date(date_str: str) -> date | None:
    """
    Converte string de data parcial em datetime.date.

    Suporta:
        - "30/7"     → date(ano_atual, 7, 30)
        - "30/07"    → date(ano_atual, 7, 30)
        - "30/7/26"  → date(2026, 7, 30)
        - "30/7/2026" → date(2026, 7, 30)
        - "30-7"     → date(ano_atual, 7, 30)
    """
    if not date_str:
        return None

    try:
        # Padroniza separadores
        date_str = date_str.replace("-", "/")
        partes = date_str.split("/")

        dia = int(partes[0])
        mes = int(partes[1])

        if len(partes) == 2:
            ano = datetime.now().year
        else:
            ano = int(partes[2])
            if ano < 100:
                ano += 2000

        return date(ano, mes, dia)
    except (ValueError, IndexError):
        return None
