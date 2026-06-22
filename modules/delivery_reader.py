"""
delivery_reader.py — Leitor do RELATÓRIO DE ENTREGAS 2026_LEVE

Estratégia de performance:
    1. Na primeira execução, o script convert_com.py converte .xlsb -> .xlsx
    2. Leituras usam o entregas_cache.xlsx (openpyxl, ~25s)
    3. No Streamlit, st.cache_data garante que só lê uma vez por sessão

Colunas principais:
    Mês, Data, Veículo, Operação, Remetente, Cliente,
    Bairro, UF, Nota_Fiscal, Peso, Volumes, Valor_Nota, Frete
"""

import os
import sys
import pandas as pd
from datetime import datetime

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from config import DELIVERIES_REPORT_PATH, FLEET_PLATES


# ══════════════════════════════════════════════════════════════
# PATHS
# ══════════════════════════════════════════════════════════════

CACHE_XLSX = os.path.join(
    os.path.dirname(os.path.abspath(DELIVERIES_REPORT_PATH)),
    "entregas_cache.xlsx",
)

# Column names we assign after skipping summary row
DELIVERY_COLUMNS = [
    "MES", "DATA", "VEICULO", "OPERACAO", "REMETENTE",
    "CLIENTE", "BAIRRO", "UF", "NOTA_FISCAL",
    "PESO", "VOLUMES", "VALOR_NOTA", "PCT_FRETE", "FRETE", "FRETE_KG",
]

COLUMNS_TO_KEEP = [
    "MES", "DATA", "VEICULO", "OPERACAO", "REMETENTE",
    "CLIENTE", "BAIRRO", "UF", "NOTA_FISCAL",
    "PESO", "VOLUMES", "VALOR_NOTA", "FRETE",
]


# ══════════════════════════════════════════════════════════════
# HELPERS
# ══════════════════════════════════════════════════════════════

def normalize_veiculo(plate_str) -> str:
    """Normaliza a coluna Veículo para uppercase."""
    if pd.isna(plate_str) or not plate_str:
        return ""
    return str(plate_str).strip().upper()


def _ensure_cache():
    """Verifica se o cache xlsx existe e é valido."""
    if os.path.exists(CACHE_XLSX):
        return True

    if not os.path.exists(DELIVERIES_REPORT_PATH):
        raise FileNotFoundError(
            f"Arquivo não encontrado: {DELIVERIES_REPORT_PATH}\n"
            f"Coloque o 'RELATÓRIO DE ENTREGAS 2026_LEVE.xlsb' na raiz do projeto."
        )

    # Try COM conversion
    try:
        import win32com.client
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        try:
            wb = excel.Workbooks.Open(
                os.path.abspath(DELIVERIES_REPORT_PATH), ReadOnly=True, UpdateLinks=0
            )
            if wb:
                wb.SaveAs(os.path.abspath(CACHE_XLSX), 51)
                wb.Close(False)
                return True
        finally:
            excel.Quit()
    except Exception:
        pass

    raise FileNotFoundError(
        f"Cache '{CACHE_XLSX}' não encontrado.\n"
        f"Execute: python scratch/convert_com.py"
    )


# ══════════════════════════════════════════════════════════════
# MAIN READER
# ══════════════════════════════════════════════════════════════

def read_deliveries_report() -> pd.DataFrame:
    """
    Lê o relatório de entregas do cache XLSX.

    Returns:
        DataFrame limpo com colunas padronizadas.
    """
    _ensure_cache()

    # Read xlsx - skip row 0 (summary) and row 1 (header "Mês", "Data"...)
    # We assign our own column names
    df = pd.read_excel(CACHE_XLSX, engine="openpyxl", skiprows=2, header=None)

    # Detect if column 0 is blank (original file) or has data (new lighter file)
    first_col_blank = df.iloc[:, 0].isna().all() or (df.iloc[:, 0].astype(str).str.strip() == "").all()
    start_col = 1 if first_col_blank else 0
    df = df.iloc[:, start_col:start_col + len(DELIVERY_COLUMNS)].copy()
    df.columns = DELIVERY_COLUMNS

    # ── Clean Data ──

    # Remove rows where MES is empty (end of data / padding)
    df = df[df["MES"].notna() & (df["MES"].astype(str).str.strip() != "")]

    # DATA already comes as datetime from xlsx (no serial conversion needed)
    df["DATA"] = pd.to_datetime(df["DATA"], errors="coerce")

    # Normalize plates
    df["VEICULO"] = df["VEICULO"].apply(normalize_veiculo)

    # Numeric columns
    for col in ["PESO", "VOLUMES", "VALOR_NOTA", "FRETE"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    if "NOTA_FISCAL" in df.columns:
        df["NOTA_FISCAL"] = pd.to_numeric(df["NOTA_FISCAL"], errors="coerce").fillna(0).astype(int)

    # String columns
    for col in ["MES", "OPERACAO", "REMETENTE", "CLIENTE", "BAIRRO", "UF"]:
        if col in df.columns:
            df[col] = df[col].fillna("").astype(str).str.strip()

    # Keep needed columns
    available = [c for c in COLUMNS_TO_KEEP if c in df.columns]
    df = df[available].copy()

    # Sort by date
    df = df.sort_values("DATA", ascending=True).reset_index(drop=True)

    return df


# ══════════════════════════════════════════════════════════════
# AGGREGATION HELPERS (for BI dashboard)
# ══════════════════════════════════════════════════════════════

def get_summary_by_vehicle(df: pd.DataFrame) -> pd.DataFrame:
    """Agrega dados por veículo."""
    return df.groupby("VEICULO").agg(
        TOTAL_NFS=("NOTA_FISCAL", "count"),
        PESO_TOTAL=("PESO", "sum"),
        VALOR_TOTAL=("VALOR_NOTA", "sum"),
        VOLUMES_TOTAL=("VOLUMES", "sum"),
        FRETE_TOTAL=("FRETE", "sum"),
    ).reset_index().sort_values("TOTAL_NFS", ascending=False)


def get_summary_by_month(df: pd.DataFrame) -> pd.DataFrame:
    """Agrega dados por mês."""
    res = df.groupby("MES").agg(
        TOTAL_NFS=("NOTA_FISCAL", "count"),
        PESO_TOTAL=("PESO", "sum"),
        VALOR_TOTAL=("VALOR_NOTA", "sum"),
        VOLUMES_TOTAL=("VOLUMES", "sum"),
    ).reset_index()
    
    # Ordenar cronologicamente
    months_order = {
        "JAN": 1, "FEV": 2, "MAR": 3, "ABR": 4, "MAI": 5, "JUN": 6,
        "JUL": 7, "AGO": 8, "SET": 9, "OUT": 10, "NOV": 11, "DEZ": 12
    }
    res["_order"] = res["MES"].str.upper().map(months_order).fillna(99)
    res = res.sort_values("_order").drop(columns=["_order"]).reset_index(drop=True)
    return res


def get_top_clients(df: pd.DataFrame, top_n: int = 10) -> pd.DataFrame:
    """Top N clientes por entregas."""
    return (
        df.groupby("CLIENTE")
        .agg(TOTAL_NFS=("NOTA_FISCAL", "count"), PESO_TOTAL=("PESO", "sum"))
        .reset_index()
        .sort_values("TOTAL_NFS", ascending=False)
        .head(top_n)
    )


def get_deliveries_by_uf(df: pd.DataFrame) -> pd.DataFrame:
    """Entregas por UF."""
    return df.groupby("UF").agg(
        TOTAL_NFS=("NOTA_FISCAL", "count"),
        VALOR_TOTAL=("VALOR_NOTA", "sum"),
    ).reset_index().sort_values("TOTAL_NFS", ascending=False)


def get_deliveries_for_plate_date(
    df: pd.DataFrame, plate: str, date: datetime
) -> pd.DataFrame:
    """NFs de um veículo em uma data (para cruzamento WhatsApp)."""
    mask = (df["VEICULO"] == plate.upper())
    if hasattr(date, 'date'):
        mask = mask & (df["DATA"].dt.date == date.date())
    return df[mask].copy()
