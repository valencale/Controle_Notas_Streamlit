"""
7_📊_BI_Analytics.py — Dashboard de BI & Analytics para entregas e viagens.

Módulo exclusivamente analítico (somente leitura):
    - KPIs de entregas e viagens
    - Gráficos interativos (Plotly)
    - Rankings de veículos e clientes

NOTA: O registro de viagens, auditoria e conferência foram migrados
para a aba 8_🚚_Expedicao.py.
"""

import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime, date
import sys, os

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from modules.ui_components import inject_custom_css, render_header
from modules.excel_handler import read_viagens

# Try to import delivery_reader (may not exist in all setups)
try:
    from modules.delivery_reader import (
        read_deliveries_report, get_summary_by_vehicle, get_summary_by_month,
        get_top_clients, get_deliveries_by_uf,
    )
    HAS_DELIVERY_REPORT = True
except ImportError:
    HAS_DELIVERY_REPORT = False

from config import FLEET_PLATES

# ══════════════════════════════════════════════════════════════
# PAGE CONFIG
# ══════════════════════════════════════════════════════════════
st.set_page_config(page_title="BI Analytics — Logística", page_icon="📊", layout="wide")
inject_custom_css()

# Plotly dark premium theme
PLOTLY_LAYOUT = dict(
    template="plotly_dark",
    paper_bgcolor="rgba(0,0,0,0)",
    plot_bgcolor="rgba(0,0,0,0)",
    font=dict(family="Inter, sans-serif", color="#cbd5e1"),
    margin=dict(l=20, r=20, t=50, b=20),
    title_font=dict(size=18, color="#f8fafc", family="Inter, sans-serif"),
    legend=dict(
        bgcolor="rgba(0,0,0,0)",
        orientation="h",
        yanchor="bottom",
        y=1.02,
        xanchor="right",
        x=1
    ),
    xaxis=dict(gridcolor="rgba(99,102,241,0.1)", zerolinecolor="rgba(99,102,241,0.2)"),
    yaxis=dict(gridcolor="rgba(99,102,241,0.1)", zerolinecolor="rgba(99,102,241,0.2)"),
    hoverlabel=dict(
        bgcolor="rgba(15, 23, 42, 0.9)",
        font_size=13,
        font_family="Inter, sans-serif",
        bordercolor="rgba(99, 102, 241, 0.4)"
    )
)

COLORS = ["#6666FF", "#33CCFF", "#00CC66", "#FFAA00", "#FF4444", "#AA00FF"]


def _hex_to_rgb(hex_color):
    h = hex_color.lstrip("#")
    if len(h) == 6:
        return f"{int(h[0:2], 16)}, {int(h[2:4], 16)}, {int(h[4:6], 16)}"
    return "99, 102, 241" # default fallback


def _bi_card(label, value, color, icon=""):
    """Colored KPI card for BI dashboard."""
    return (
        '<div style="'
        f'background: linear-gradient(135deg, rgba({_hex_to_rgb(color)}, 0.15), rgba({_hex_to_rgb(color)}, 0.05));'
        f'border: 1px solid rgba({_hex_to_rgb(color)}, 0.25);'
        f'border-left: 4px solid {color};'
        'border-radius: 14px;'
        'padding: 16px 20px;'
        'box-shadow: 0 4px 16px rgba(0,0,0,0.2);'
        'transition: transform 0.2s;'
        '">'
        f'<div style="color: {color}; font-size: 12px; font-weight: 700;'
        f' text-transform: uppercase; letter-spacing: 0.8px; margin-bottom: 4px;">'
        f'{icon} {label}</div>'
        f'<div style="color: #f8fafc; font-size: 28px; font-weight: 700;">'
        f'{value}</div>'
        '</div>'
    )


# ══════════════════════════════════════════════════════════════
# DATA LOADING (cached)
# ══════════════════════════════════════════════════════════════
@st.cache_data(show_spinner="Carregando relatório de entregas...", ttl=300)
def load_deliveries():
    if not HAS_DELIVERY_REPORT:
        return pd.DataFrame()
    try:
        return read_deliveries_report()
    except FileNotFoundError:
        return pd.DataFrame()


@st.cache_data(show_spinner="Carregando viagens...", ttl=300)
def load_viagens():
    return read_viagens()


# ══════════════════════════════════════════════════════════════
# HEADER
# ══════════════════════════════════════════════════════════════
render_header("BI & Analytics", "Inteligência operacional sobre entregas e viagens")

st.caption("📌 Para registrar viagens, auditoria e conferência, acesse a aba **🚚 Expedição**.")

# ══════════════════════════════════════════════════════════════
# DASHBOARD
# ══════════════════════════════════════════════════════════════
df_del = load_deliveries()
df_viag = load_viagens()

if df_del.empty:
    st.warning(
        "Relatório de entregas não encontrado. "
        "Execute `python scratch/convert_com.py` para gerar o cache."
    )
    st.stop()

# ── Filters ──
with st.sidebar:
    st.markdown("### 🔧 Filtros BI")
    meses = sorted(df_del["MES"].unique().tolist())
    sel_mes = st.multiselect("Mês", meses, default=meses, key="bi_mes")

    veiculos = sorted(df_del["VEICULO"].unique().tolist())
    sel_veic = st.multiselect("Veículo", veiculos, default=FLEET_PLATES, key="bi_veic")

    remetentes = sorted(df_del["REMETENTE"].unique().tolist())
    sel_rem = st.multiselect("Remetente", remetentes, default=remetentes, key="bi_rem")

# Apply filters
mask = (
    df_del["MES"].isin(sel_mes) &
    df_del["VEICULO"].isin(sel_veic) &
    df_del["REMETENTE"].isin(sel_rem)
)
df_f = df_del[mask].copy()

# ── KPIs Row 1 — Relatório de Entregas ──
st.markdown('<div style="color: #94a3b8; font-size: 12px; font-weight: 600; text-transform: uppercase; letter-spacing: 1px; margin-bottom: 8px;">📦 Relatório de Entregas</div>', unsafe_allow_html=True)
total_nfs = len(df_f)
peso_total = df_f["PESO"].sum()
valor_total = df_f["VALOR_NOTA"].sum()
volumes_total = df_f["VOLUMES"].sum()

k1, k2, k3, k4 = st.columns(4)
with k1:
    st.markdown(_bi_card("Total NFs", f"{total_nfs:,}".replace(",", "."), "#6666FF", "📄"), unsafe_allow_html=True)
with k2:
    st.markdown(_bi_card("Peso Total", f"{peso_total:,.0f} kg".replace(",", "."), "#33CCFF", "⚖️"), unsafe_allow_html=True)
with k3:
    st.markdown(_bi_card("Valor Transportado", f"R$ {valor_total:,.0f}".replace(",", "."), "#00CC66", "💰"), unsafe_allow_html=True)
with k4:
    st.markdown(_bi_card("Volumes", f"{volumes_total:,.0f}".replace(",", "."), "#FFAA00", "📦"), unsafe_allow_html=True)

# ── KPIs Row 2 — Viagens WhatsApp ──
if not df_viag.empty:
    st.markdown('<div style="color: #94a3b8; font-size: 12px; font-weight: 600; text-transform: uppercase; letter-spacing: 1px; margin: 24px 0 8px 0;">🚚 Viagens (WhatsApp)</div>', unsafe_allow_html=True)
    total_viagens = len(df_viag)
    km_total = pd.to_numeric(df_viag.get("KM_RODADO", 0), errors="coerce").sum()
    total_entregas = pd.to_numeric(df_viag.get("ENTREGAS", 0), errors="coerce").sum()
    total_coletas = pd.to_numeric(df_viag.get("COLETAS", 0), errors="coerce").sum()
    total_pedidos = pd.to_numeric(df_viag.get("PEDIDOS", 0), errors="coerce").sum()
    taxa_sucesso = (total_entregas / total_pedidos * 100) if total_pedidos > 0 else 0

    v1, v2, v3, v4 = st.columns(4)
    with v1:
        st.markdown(_bi_card("Viagens", str(int(total_viagens)), "#AA00FF", "🚛"), unsafe_allow_html=True)
    with v2:
        st.markdown(_bi_card("KM Rodado", f"{km_total:,.0f} km".replace(",", "."), "#33CCFF", "🛣️"), unsafe_allow_html=True)
    with v3:
        color_taxa = "#00CC66" if taxa_sucesso >= 85 else "#FFAA00" if taxa_sucesso >= 70 else "#FF4444"
        st.markdown(_bi_card("Taxa Sucesso", f"{taxa_sucesso:.1f}%", color_taxa, "📈"), unsafe_allow_html=True)
    with v4:
        st.markdown(_bi_card("Coletas", str(int(total_coletas)), "#FF4444", "📥"), unsafe_allow_html=True)

st.markdown("---")

# ── Charts ──
c1, c2 = st.columns(2)

with c1:
    sv = get_summary_by_vehicle(df_f)
    sv_fleet = sv[sv["VEICULO"].isin(FLEET_PLATES)]
    if not sv_fleet.empty:
        fig = px.bar(
            sv_fleet, x="VEICULO", y="TOTAL_NFS",
            color="VEICULO", color_discrete_sequence=COLORS,
            title="Entregas por Veículo (Frota)",
            labels={"TOTAL_NFS": "Total NFs", "VEICULO": "Placa do Veículo"},
            text="TOTAL_NFS"
        )
        fig.update_traces(textposition='outside', hovertemplate="<b>Placa:</b> %{x}<br><b>Total NFs:</b> %{y}<extra></extra>")
        fig.update_layout(**PLOTLY_LAYOUT)
        st.plotly_chart(fig, use_container_width=True)

with c2:
    tc = get_top_clients(df_f, 10)
    if not tc.empty:
        fig = px.bar(
            tc, x="TOTAL_NFS", y="CLIENTE", orientation="h",
            color="TOTAL_NFS", color_continuous_scale="Plasma",
            title="Top 10 Clientes por Entregas",
            labels={"TOTAL_NFS": "Total de Entregas", "CLIENTE": "Nome do Cliente"}
        )
        fig.update_traces(hovertemplate="<b>Cliente:</b> %{y}<br><b>Entregas:</b> %{x}<extra></extra>")
        fig.update_layout(**PLOTLY_LAYOUT, height=400)
        st.plotly_chart(fig, use_container_width=True)

c3, c4 = st.columns(2)

with c3:
    sm = get_summary_by_month(df_f)
    if not sm.empty:
        fig = px.line(
            sm, x="MES", y="TOTAL_NFS",
            markers=True, color_discrete_sequence=["#33CCFF"],
            title="Evolução Mensal de NFs",
            labels={"TOTAL_NFS": "Total NFs", "MES": "Mês de Entrega"}
        )
        fig.update_traces(line=dict(width=3), marker=dict(size=8), hovertemplate="<b>Mês:</b> %{x}<br><b>NFs:</b> %{y}<extra></extra>")
        fig.update_layout(**PLOTLY_LAYOUT)
        fig.update_xaxes(categoryorder='array', categoryarray=sm["MES"])
        st.plotly_chart(fig, use_container_width=True)

with c4:
    uf = get_deliveries_by_uf(df_f)
    if not uf.empty and len(uf) > 1:
        fig = px.treemap(
            uf, path=["UF"], values="TOTAL_NFS",
            color="VALOR_TOTAL", color_continuous_scale="Viridis",
            title="Distribuição Geográfica (UF)",
        )
        fig.update_traces(hovertemplate="<b>UF:</b> %{label}<br><b>NFs:</b> %{value}<br><b>Valor Total:</b> R$ %{color:,.2f}<extra></extra>")
        fig.update_layout(**PLOTLY_LAYOUT, height=400)
        st.plotly_chart(fig, use_container_width=True)

# Distribuição de operações
if not df_f.empty:
    ops = df_f["OPERACAO"].value_counts().reset_index()
    ops.columns = ["OPERACAO", "TOTAL"]
    fig = px.pie(
        ops, names="OPERACAO", values="TOTAL",
        color_discrete_sequence=COLORS,
        title="Distribuição por Tipo de Operação",
        hole=0.5,
    )
    fig.update_traces(hovertemplate="<b>Operação:</b> %{label}<br><b>Total:</b> %{value} (%{percent})<extra></extra>")
    fig.update_layout(**PLOTLY_LAYOUT, height=350)
    st.plotly_chart(fig, use_container_width=True)
