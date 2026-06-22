"""
config.py — Constantes centralizadas para a aplicação de gestão logística.

Define caminhos, nomes de abas, colunas, status válidos e parâmetros
de configuração global usados por todos os módulos.
"""

import os

# ══════════════════════════════════════════════════════════════
# PATHS
# ══════════════════════════════════════════════════════════════
EXCEL_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "CONTROLE NOTAS.xlsm")

# ══════════════════════════════════════════════════════════════
# SHEET NAMES
# ══════════════════════════════════════════════════════════════
SHEET_PRINCIPAL = "Dados"
SHEET_HISTORICO = "Historico"
SHEET_SETUP = "Setup"

# ══════════════════════════════════════════════════════════════
# DATA STRUCTURE — Aba "Dados"
# Row 10 = Header row
# Rows 1-6 = Summary/KPI area (columns E-F)
# Rows 7-9 = Spacing / title
# Row 11+ = Data rows
# ══════════════════════════════════════════════════════════════
HEADER_ROW = 10
DATA_START_ROW = 11

# Column mapping (1-indexed as openpyxl uses)
COLUMNS = {
    "DATA": 2,       # B
    "CLIENTE": 3,    # C
    "PEDIDO": 4,     # D
    "EMPRESA": 5,    # E
    "NF": 6,         # F — Nota Fiscal
    "STATUS": 7,     # G
    "OBS": 8,        # H
    "ENDERECO": 9,   # I — Endereço (extraído de PDFs)
}

# Column headers in order (for DataFrame conversion)
COLUMN_NAMES = ["DATA", "CLIENTE", "PEDIDO", "EMPRESA", "NF", "STATUS", "OBS", "ENDERECO"]

# Excel header display names (internal key → Excel header text)
EXCEL_HEADER_NAMES = {
    "DATA": "DATA",
    "CLIENTE": "CLIENTE",
    "PEDIDO": "PEDIDO",
    "EMPRESA": "EMPRESA",
    "NF": "NOTA FISCAL",
    "STATUS": "STATUS",
    "OBS": "OBS",
    "ENDERECO": "ENDERECO",
}

# Primary Key column
PK_COLUMN = "PEDIDO"
PK_COL_INDEX = COLUMNS["PEDIDO"]  # Column D = 4

# ══════════════════════════════════════════════════════════════
# STATUS OPTIONS
# ══════════════════════════════════════════════════════════════
STATUS_OPTIONS = [
    "SEPARACAO",
    "PARCIAL",
    "AUSENTE",
    "CONCLUIDO",
    "AGUARDANDO NF",
    "SEM MATERIAL",
    "ENVIAR DATA",
    "EM ROTA",
    "ENTREGUE",
]

# Status colors for UI
STATUS_COLORS = {
    "SEPARACAO": "#FFD700",       # Gold
    "PARCIAL": "#FF8C00",         # Dark Orange
    "AUSENTE": "#DC143C",         # Crimson
    "CONCLUIDO": "#32CD32",       # Lime Green
    "AGUARDANDO NF": "#FF69B4",   # Hot Pink
    "SEM MATERIAL": "#E74C3C",    # Red-Orange
    "ENVIAR DATA": "#E040FB",     # Purple/Magenta
    "EM ROTA": "#3B82F6",         # Blue
    "ENTREGUE": "#059669",        # Emerald
}

# Status icons for map markers and UI
STATUS_ICONS = {
    "SEPARACAO": "📦",
    "PARCIAL": "⚠️",
    "AUSENTE": "❌",
    "CONCLUIDO": "✅",
    "AGUARDANDO NF": "📄",
    "SEM MATERIAL": "🚫",
    "ENVIAR DATA": "📤",
    "EM ROTA": "🚚",
    "ENTREGUE": "🏁",
}

# ══════════════════════════════════════════════════════════════
# EMPRESA (Carrier) OPTIONS
# ══════════════════════════════════════════════════════════════
EMPRESA_OPTIONS = [
    "GREEN BAGS",
    "ONFINITY",
]

# ══════════════════════════════════════════════════════════════
# MAP / GEOCODING
# ══════════════════════════════════════════════════════════════
MAP_CENTER = [-23.5325, -46.7922]  # Osasco, SP
MAP_ZOOM = 11
NOMINATIM_USER_AGENT = "streamlit-logistica-app"
GEOCODE_TIMEOUT = 5  # seconds

# ══════════════════════════════════════════════════════════════
# PDF PARSING — Onfinity "Mapa de Separação por Pedido"
# ══════════════════════════════════════════════════════════════
PDF_DEFAULT_STATUS = "SEPARACAO"
PDF_DEFAULT_EMPRESA = "ONFINITY"

# ══════════════════════════════════════════════════════════════
# BI & ANALYTICS — Relatório de Entregas + Viagens WhatsApp
# ══════════════════════════════════════════════════════════════

# Caminho do relatório de entregas (arquivo externo .xlsb)
DELIVERIES_REPORT_PATH = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    "RELATÓRIO DE ENTREGAS 2026_LEVE.xlsb",
)
DELIVERIES_SHEET = "ENTREGAS-26"
DELIVERIES_HEADER_ROW = 1  # Row 1 = summary area, Row 2 = headers

# Caminho do arquivo de viagens WhatsApp (arquivo separado, sem macros)
VIAGENS_PATH = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    "viagens_whatsapp.xlsx",
)

# Colunas do arquivo de viagens
VIAGENS_COLUMNS = [
    "DATA", "MOTORISTA", "AJUDANTE", "PLACA",
    "PEDIDOS", "ENTREGAS", "COLETAS",
    "KM_INICIAL", "KM_FINAL", "KM_RODADO",
    "OCORRENCIAS", "OBS", "TEXTO_ORIGINAL",
]

# Frota conhecida — 3 veículos principais
# Mapa de normalização: variações comuns → placa canônica
KNOWN_PLATES = {
    "QJJ-9302": "QJJ-9302",
    "QJJ9302": "QJJ-9302",
    "QJJ 9302": "QJJ-9302",
    "RLK-0E24": "RLK-0E24",
    "RLK0E24": "RLK-0E24",
    "RLK 0E24": "RLK-0E24",
    "RLK 0e24": "RLK-0E24",
    "Rlk 0e24": "RLK-0E24",
    "KZG-9D54": "KZG-9D54",
    "KZG9D54": "KZG-9D54",
    "Kzg9d54": "KZG-9D54",
    "KZG 9D54": "KZG-9D54",
}

FLEET_PLATES = ["QJJ-9302", "RLK-0E24", "KZG-9D54"]

# ══════════════════════════════════════════════════════════════
# EXPEDIÇÃO DIÁRIA — Planejamento de saída de veículos
# ══════════════════════════════════════════════════════════════

# Caminho do arquivo de expedições (separado do .xlsm para segurança)
EXPEDITIONS_PATH = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    "expedicoes.xlsx",
)

# Colunas da aba "Expedições"
EXPEDITION_COLUMNS = [
    "DATA_EXPEDICAO", "PEDIDO", "CLIENTE", "DESTINO",
    "VEICULO", "MOTORISTA", "ORDEM_ENTREGA",
    "NF", "PESO", "VOLUMES", "CARREGADO", "OBS", "CRIADO_EM",
]

# Colunas da aba "Motoristas" (configuração padrão por veículo)
MOTORISTAS_COLUMNS = ["PLACA", "MOTORISTA_PADRAO"]

# Motoristas padrão — ajustar conforme necessidade
DEFAULT_DRIVERS = {
    "QJJ-9302": "Jefferson",
    "RLK-0E24": "Jose",
    "KZG-9D54": "Renato",
}
