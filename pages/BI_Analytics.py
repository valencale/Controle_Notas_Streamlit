"""
BI_Analytics.py — Dashboard de BI & Analytics avançado.

Seções:
    1. KPIs por tipo de operação (Entrada, Saída, Devolução, etc.)
    2. Evolução semanal e mensal de notas fiscais (barras)
    3. Pareto de remetentes (logística avançada)
    4. Análise de anomalias nos valores das notas (scatter)
    5. Visão por clientes (top clientes, distribuição de peso/valor)
    6. Viagens WhatsApp KPIs
    7. Distribuição geográfica (UF)
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, date, timedelta
import sys, os

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from modules.ui_components import inject_custom_css, render_header
from modules.excel_handler import read_viagens, read_principal

# Try to import delivery_reader (may not exist in all setups)
try:
    from modules.delivery_reader import (
        read_deliveries_report, get_summary_by_vehicle, get_summary_by_month,
        get_top_clients, get_deliveries_by_uf,
    )
    HAS_DELIVERY_REPORT = True
except ImportError:
    HAS_DELIVERY_REPORT = False

from config import FLEET_PLATES, STATUS_COLORS, STATUS_OPTIONS

# ══════════════════════════════════════════════════════════════
# PAGE CONFIG
# ══════════════════════════════════════════════════════════════
inject_custom_css()

# ── Plotly dark premium theme ──
PLOTLY_LAYOUT = dict(
    template="plotly_dark",
    paper_bgcolor="rgba(0,0,0,0)",
    plot_bgcolor="rgba(0,0,0,0)",
    font=dict(family="Inter, sans-serif", color="#cbd5e1", size=13),
    margin=dict(l=24, r=24, t=56, b=24),
    title_font=dict(size=17, color="#f8fafc", family="Inter, sans-serif"),
    legend=dict(
        bgcolor="rgba(0,0,0,0)",
        orientation="h",
        yanchor="bottom",
        y=1.02,
        xanchor="right",
        x=1
    ),
    xaxis=dict(
        gridcolor="rgba(16,185,129,0.08)",
        zerolinecolor="rgba(16,185,129,0.15)",
    ),
    yaxis=dict(
        gridcolor="rgba(16,185,129,0.08)",
        zerolinecolor="rgba(16,185,129,0.15)",
    ),
    hoverlabel=dict(
        bgcolor="rgba(15, 23, 42, 0.95)",
        font_size=13,
        font_family="Inter, sans-serif",
        bordercolor="rgba(16, 185, 129, 0.4)",
    ),
    bargap=0.2,
)

# Paleta harmonizada com o tema Green Bags
COLORS = ["#10b981", "#06d6a0", "#33CCFF", "#6366f1", "#f59e0b", "#ef4444", "#8b5cf6", "#ec4899"]
COLOR_ACCENT = "#10b981"
COLOR_WARN = "#f59e0b"
COLOR_DANGER = "#ef4444"
COLOR_INFO = "#33CCFF"


def _hex_to_rgb(hex_color):
    h = hex_color.lstrip("#")
    if len(h) == 6:
        return f"{int(h[0:2], 16)}, {int(h[2:4], 16)}, {int(h[4:6], 16)}"
    return "16, 185, 129"


def _bi_card(label, value, color, icon="", delta=None):
    """Premium KPI card with optional delta indicator."""
    delta_html = ""
    if delta is not None:
        arrow = "↑" if delta >= 0 else "↓"
        delta_color = "#10b981" if delta >= 0 else "#ef4444"
        delta_html = (
            f'<div style="font-size: 12px; font-weight: 600; color: {delta_color}; '
            f'margin-top: 4px;">{arrow} {abs(delta):.1f}%</div>'
        )
    return (
        '<div style="'
        f'background: linear-gradient(135deg, rgba({_hex_to_rgb(color)}, 0.12), rgba({_hex_to_rgb(color)}, 0.04));'
        f'border: 1px solid rgba({_hex_to_rgb(color)}, 0.2);'
        f'border-left: 4px solid {color};'
        'border-radius: 14px;'
        'padding: 18px 22px;'
        'box-shadow: 0 4px 16px rgba(0,0,0,0.15);'
        'transition: transform 0.2s;'
        '">'
        f'<div style="color: {color}; font-size: 11px; font-weight: 700;'
        f' text-transform: uppercase; letter-spacing: 0.8px; margin-bottom: 6px;">'
        f'{icon} {label}</div>'
        f'<div style="color: #f8fafc; font-size: 28px; font-weight: 700; line-height: 1.2;">'
        f'{value}</div>'
        f'{delta_html}'
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


@st.cache_data(show_spinner="Carregando pedidos...", ttl=300)
def load_pedidos():
    try:
        return read_principal()
    except Exception:
        return pd.DataFrame()


# ══════════════════════════════════════════════════════════════
# HEADER
# ══════════════════════════════════════════════════════════════
render_header("BI & Analytics", "Inteligência operacional sobre entregas, viagens e clientes")

# ══════════════════════════════════════════════════════════════
# DATA
# ══════════════════════════════════════════════════════════════
df_del = load_deliveries()
df_viag = load_viagens()
df_pedidos = load_pedidos()

if df_del.empty:
    st.warning(
        "Relatório de entregas não encontrado. "
        "Clique em **🔄 Atualizar Cache Entregas** na sidebar de Administração."
    )
    st.stop()

# ══════════════════════════════════════════════════════════════
# FILTERS (sidebar)
# ══════════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown("### 🔧 Filtros BI")

    # Período por data
    if "DATA" in df_del.columns and df_del["DATA"].notna().any():
        min_date = df_del["DATA"].min().date()
        max_date = df_del["DATA"].max().date()
        periodo = st.date_input(
            "📅 Período",
            value=(min_date, max_date),
            min_value=min_date,
            max_value=max_date,
            key="bi_periodo",
        )
    else:
        periodo = None

    # Remetente
    remetentes = sorted(df_del["REMETENTE"].unique().tolist())
    sel_rem = st.multiselect("📦 Remetente", remetentes, default=remetentes, key="bi_rem")

    # Veículo
    veiculos = sorted(df_del["VEICULO"].unique().tolist())
    sel_veic = st.multiselect("🚛 Veículo", veiculos, default=FLEET_PLATES, key="bi_veic")

    # Operação
    operacoes = sorted(df_del["OPERACAO"].unique().tolist())
    sel_op = st.multiselect("⚙️ Operação", operacoes, default=operacoes, key="bi_op")

    # Cliente (filtro extra)
    clientes_all = sorted(df_del["CLIENTE"].unique().tolist())
    sel_cli = st.multiselect("👤 Cliente", clientes_all, default=[], key="bi_cli",
                              help="Deixe vazio para ver todos")

# ── Aplicar filtros ──
mask = (
    df_del["REMETENTE"].isin(sel_rem) &
    df_del["VEICULO"].isin(sel_veic) &
    df_del["OPERACAO"].isin(sel_op)
)

if periodo and len(periodo) == 2:
    mask = mask & (df_del["DATA"].dt.date >= periodo[0]) & (df_del["DATA"].dt.date <= periodo[1])

if sel_cli:
    mask = mask & df_del["CLIENTE"].isin(sel_cli)

df_f = df_del[mask].copy()

if df_f.empty:
    st.info("Nenhum dado encontrado para os filtros selecionados.")
    st.stop()


# ══════════════════════════════════════════════════════════════
# 1. KPIs POR TIPO DE OPERAÇÃO
# ══════════════════════════════════════════════════════════════
st.markdown(
    '<div style="color: #10b981; font-size: 13px; font-weight: 700; '
    'text-transform: uppercase; letter-spacing: 1px; margin-bottom: 12px;">'
    '📊 Resumo por Tipo de Operação</div>',
    unsafe_allow_html=True,
)

# KPIs globais
total_nfs = len(df_f)
peso_total = df_f["PESO"].sum()
valor_total = df_f["VALOR_NOTA"].sum()
volumes_total = df_f["VOLUMES"].sum()

k1, k2, k3, k4 = st.columns(4)
with k1:
    st.markdown(_bi_card("Total NFs", f"{total_nfs:,}".replace(",", "."), "#10b981", "📄"), unsafe_allow_html=True)
with k2:
    st.markdown(_bi_card("Peso Total", f"{peso_total:,.0f} kg".replace(",", "."), "#33CCFF", "⚖️"), unsafe_allow_html=True)
with k3:
    st.markdown(_bi_card("Valor Transportado", f"R$ {valor_total:,.0f}".replace(",", "."), "#6366f1", "💰"), unsafe_allow_html=True)
with k4:
    st.markdown(_bi_card("Volumes", f"{volumes_total:,.0f}".replace(",", "."), "#f59e0b", "📦"), unsafe_allow_html=True)

# KPIs por operação
ops = df_f["OPERACAO"].value_counts()
if len(ops) > 0:
    op_colors = {"ENTRADA": "#10b981", "SAÍDA": "#33CCFF", "SAIDA": "#33CCFF",
                 "DEVOLUÇÃO": "#ef4444", "DEVOLUCAO": "#ef4444",
                 "TRANSFERÊNCIA": "#8b5cf6", "TRANSFERENCIA": "#8b5cf6"}
    op_cols = st.columns(min(len(ops), 6))
    for i, (op_name, count) in enumerate(ops.items()):
        col_idx = i % len(op_cols)
        color = op_colors.get(op_name.upper(), COLORS[i % len(COLORS)])
        with op_cols[col_idx]:
            pct = count / total_nfs * 100
            st.markdown(
                _bi_card(op_name, f"{count:,}".replace(",", "."), color, "",
                         delta=None),
                unsafe_allow_html=True,
            )

st.markdown("---")

# ══════════════════════════════════════════════════════════════
# 2. EVOLUÇÃO DE NOTAS: SEMANAL + MENSAL
# ══════════════════════════════════════════════════════════════
st.markdown(
    '<div style="color: #10b981; font-size: 13px; font-weight: 700; '
    'text-transform: uppercase; letter-spacing: 1px; margin-bottom: 12px;">'
    '📈 Evolução de Notas Fiscais</div>',
    unsafe_allow_html=True,
)

c_week, c_month = st.columns(2)

# ── Gráfico Semanal ──
with c_week:
    df_week = df_f.copy()
    df_week["SEMANA"] = df_week["DATA"].dt.to_period("W").apply(lambda r: r.start_time)
    weekly = df_week.groupby("SEMANA").agg(
        TOTAL_NFS=("NOTA_FISCAL", "count"),
        PESO=("PESO", "sum"),
        VALOR=("VALOR_NOTA", "sum"),
    ).reset_index().sort_values("SEMANA")

    if not weekly.empty:
        weekly["SEMANA_STR"] = weekly["SEMANA"].dt.strftime("%d/%m")
        fig = px.bar(
            weekly, x="SEMANA_STR", y="TOTAL_NFS",
            color_discrete_sequence=[COLOR_ACCENT],
            title="Notas por Semana",
            labels={"TOTAL_NFS": "NFs", "SEMANA_STR": "Semana (início)"},
            text="TOTAL_NFS",
        )
        fig.update_traces(
            textposition="outside",
            hovertemplate="<b>Semana:</b> %{x}<br><b>NFs:</b> %{y}<extra></extra>",
            marker=dict(cornerradius=6),
        )
        fig.update_layout(**PLOTLY_LAYOUT, height=400)
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("Sem dados semanais para exibir.")

# ── Gráfico Mensal ──
with c_month:
    sm = get_summary_by_month(df_f)
    if not sm.empty:
        fig = px.bar(
            sm, x="MES", y="TOTAL_NFS",
            color_discrete_sequence=["#6366f1"],
            title="Notas por Mês",
            labels={"TOTAL_NFS": "NFs", "MES": "Mês"},
            text="TOTAL_NFS",
        )
        fig.update_traces(
            textposition="outside",
            hovertemplate="<b>Mês:</b> %{x}<br><b>NFs:</b> %{y}<extra></extra>",
            marker=dict(cornerradius=6),
        )
        fig.update_xaxes(categoryorder='array', categoryarray=sm["MES"])
        fig.update_layout(**PLOTLY_LAYOUT, height=400)
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("Sem dados mensais para exibir.")

st.markdown("---")

# ══════════════════════════════════════════════════════════════
# 3. LOGÍSTICA AVANÇADA: PARETO DE REMETENTES + ANOMALIAS
# ══════════════════════════════════════════════════════════════
st.markdown(
    '<div style="color: #10b981; font-size: 13px; font-weight: 700; '
    'text-transform: uppercase; letter-spacing: 1px; margin-bottom: 12px;">'
    '🔬 Logística Avançada</div>',
    unsafe_allow_html=True,
)

c_pareto, c_anomaly = st.columns(2)

# ── Pareto de Remetentes ──
with c_pareto:
    pareto_metric = st.radio("Métrica do Pareto:", ["Quantidade (NFs)", "Valor (R$)"], horizontal=True, key="pareto_radio")
    
    rem_agg = (
        df_f.groupby("REMETENTE")
        .agg(TOTAL_NFS=("NOTA_FISCAL", "count"), VALOR=("VALOR_NOTA", "sum"))
        .reset_index()
    )
    
    sort_col = "TOTAL_NFS" if pareto_metric == "Quantidade (NFs)" else "VALOR"
    y_title = "Total NFs" if pareto_metric == "Quantidade (NFs)" else "Valor (R$)"
    y_format = "NFs: %{y}" if pareto_metric == "Quantidade (NFs)" else "Valor: R$ %{y:,.2f}"
    
    rem_agg = rem_agg.sort_values(sort_col, ascending=False)
    
    if not rem_agg.empty and rem_agg[sort_col].sum() > 0:
        rem_agg["PCT_ACUM"] = rem_agg[sort_col].cumsum() / rem_agg[sort_col].sum() * 100

        fig = go.Figure()
        fig.add_trace(go.Bar(
            x=rem_agg["REMETENTE"], y=rem_agg[sort_col],
            name=y_title,
            marker_color=COLOR_ACCENT,
            marker_cornerradius=6,
            hovertemplate=f"<b>%{{x}}</b><br>{y_format}<extra></extra>",
        ))
        fig.add_trace(go.Scatter(
            x=rem_agg["REMETENTE"], y=rem_agg["PCT_ACUM"],
            name="% Acumulado",
            yaxis="y2",
            mode="lines+markers",
            line=dict(color="#f59e0b", width=2.5),
            marker=dict(size=6),
            hovertemplate="<b>%{x}</b><br>Acumulado: %{y:.1f}%<extra></extra>",
        ))
        fig.add_hline(y=80, yref="y2", line_dash="dash",
                      line_color="rgba(239,68,68,0.5)", annotation_text="80%",
                      annotation_position="top right")
        pareto_layout = {k: v for k, v in PLOTLY_LAYOUT.items() if k not in ("yaxis",)}
        fig.update_layout(
            **pareto_layout,
            title=f"Pareto — {y_title}",
            yaxis=dict(title=y_title, gridcolor="rgba(16,185,129,0.08)"),
            yaxis2=dict(title="% Acumulado", overlaying="y", side="right",
                        range=[0, 105], gridcolor="rgba(0,0,0,0)"),
            height=420,
            showlegend=True,
        )
        st.plotly_chart(fig, use_container_width=True)

# ── Anomalias de Valores ──
with c_anomaly:
    df_scatter = df_f[df_f["VALOR_NOTA"] > 0].copy()
    if not df_scatter.empty and len(df_scatter) > 5:
        # Calcula Z-score para detectar anomalias
        mean_val = df_scatter["VALOR_NOTA"].mean()
        std_val = df_scatter["VALOR_NOTA"].std()
        if std_val > 0:
            df_scatter["Z_SCORE"] = (df_scatter["VALOR_NOTA"] - mean_val) / std_val
            df_scatter["ANOMALIA"] = df_scatter["Z_SCORE"].abs() > 2
        else:
            df_scatter["Z_SCORE"] = 0
            df_scatter["ANOMALIA"] = False

        fig = px.scatter(
            df_scatter, x="PESO", y="VALOR_NOTA",
            color="ANOMALIA",
            color_discrete_map={True: COLOR_DANGER, False: "rgba(16,185,129,0.5)"},
            title="Detecção de Anomalias (Peso × Valor)",
            labels={"PESO": "Peso (kg)", "VALOR_NOTA": "Valor da Nota (R$)", "ANOMALIA": "Anomalia"},
            hover_data=["CLIENTE", "REMETENTE"],
            opacity=0.7,
        )
        fig.update_traces(
            marker=dict(size=6, line=dict(width=0)),
            hovertemplate=(
                "<b>Cliente:</b> %{customdata[0]}<br>"
                "<b>Remetente:</b> %{customdata[1]}<br>"
                "<b>Peso:</b> %{x:,.0f} kg<br>"
                "<b>Valor:</b> R$ %{y:,.2f}<extra></extra>"
            ),
        )
        fig.update_layout(**PLOTLY_LAYOUT, height=420)
        st.plotly_chart(fig, use_container_width=True)

        anomalies = df_scatter[df_scatter["ANOMALIA"]]
        if not anomalies.empty:
            st.caption(f"⚠️ {len(anomalies)} notas com valor fora do padrão (Z-score > 2)")
    else:
        st.info("Dados insuficientes para análise de anomalias.")

st.markdown("---")

# ══════════════════════════════════════════════════════════════
# 4. VISÃO POR CLIENTES
# ══════════════════════════════════════════════════════════════
st.markdown(
    '<div style="color: #10b981; font-size: 13px; font-weight: 700; '
    'text-transform: uppercase; letter-spacing: 1px; margin-bottom: 12px;">'
    '👤 Visão por Clientes</div>',
    unsafe_allow_html=True,
)

c_top, c_dist = st.columns(2)

with c_top:
    tc = get_top_clients(df_f, 15)
    if not tc.empty:
        fig = px.bar(
            tc, x="TOTAL_NFS", y="CLIENTE", orientation="h",
            color="TOTAL_NFS",
            color_continuous_scale=[[0, "#064e3b"], [0.5, "#10b981"], [1, "#6ee7b7"]],
            title="Top 15 Clientes por Entregas",
            labels={"TOTAL_NFS": "Total de Entregas", "CLIENTE": ""},
            text="TOTAL_NFS",
        )
        fig.update_traces(
            textposition="outside",
            hovertemplate="<b>%{y}</b><br>Entregas: %{x}<extra></extra>",
        )
        fig.update_layout(**PLOTLY_LAYOUT, height=520, coloraxis_showscale=False)
        fig.update_yaxes(categoryorder="total ascending")
        st.plotly_chart(fig, use_container_width=True)

with c_dist:
    # Distribuição de valor por cliente (top 10)
    cli_val = (
        df_f.groupby("CLIENTE")
        .agg(VALOR_TOTAL=("VALOR_NOTA", "sum"), PESO_TOTAL=("PESO", "sum"),
             NFS=("NOTA_FISCAL", "count"))
        .reset_index()
        .sort_values("VALOR_TOTAL", ascending=False)
        .head(10)
    )
    if not cli_val.empty:
        fig = px.bar(
            cli_val, x="CLIENTE", y="VALOR_TOTAL",
            color="NFS",
            color_continuous_scale=[[0, "#312e81"], [0.5, "#6366f1"], [1, "#a5b4fc"]],
            title="Top 10 Clientes por Valor Transportado",
            labels={"VALOR_TOTAL": "Valor (R$)", "CLIENTE": "", "NFS": "NFs"},
            text=cli_val["VALOR_TOTAL"].apply(lambda v: f"R$ {v:,.0f}".replace(",", ".")),
        )
        fig.update_traces(
            textposition="outside",
            hovertemplate="<b>%{x}</b><br>Valor: R$ %{y:,.2f}<br>NFs: %{marker.color}<extra></extra>",
            marker_cornerradius=6,
        )
        fig.update_layout(**PLOTLY_LAYOUT, height=520, coloraxis_showscale=False)
        fig.update_xaxes(tickangle=-45)
        st.plotly_chart(fig, use_container_width=True)

# ── Tabela de clientes expandível ──
with st.expander("📋 Tabela completa de clientes"):
    cli_full = (
        df_f.groupby("CLIENTE")
        .agg(
            NFs=("NOTA_FISCAL", "count"),
            Peso_kg=("PESO", "sum"),
            Volumes=("VOLUMES", "sum"),
            Valor_R=("VALOR_NOTA", "sum"),
            Frete_R=("FRETE", "sum"),
        )
        .reset_index()
        .sort_values("NFs", ascending=False)
    )
    cli_full["Ticket Médio"] = (cli_full["Valor_R"] / cli_full["NFs"]).round(2)
    st.dataframe(
        cli_full.style.format({
            "Peso_kg": "{:,.0f}",
            "Volumes": "{:,.0f}",
            "Valor_R": "R$ {:,.2f}",
            "Frete_R": "R$ {:,.2f}",
            "Ticket Médio": "R$ {:,.2f}",
        }),
        use_container_width=True,
        height=400,
    )

st.markdown("---")

# ══════════════════════════════════════════════════════════════
# 5. VIAGENS WHATSAPP
# ══════════════════════════════════════════════════════════════
if not df_viag.empty:
    st.markdown(
        '<div style="color: #10b981; font-size: 13px; font-weight: 700; '
        'text-transform: uppercase; letter-spacing: 1px; margin-bottom: 12px;">'
        '🚚 Viagens (WhatsApp)</div>',
        unsafe_allow_html=True,
    )

    total_viagens = len(df_viag)
    km_total = pd.to_numeric(df_viag.get("KM_RODADO", 0), errors="coerce").sum()
    total_entregas = pd.to_numeric(df_viag.get("ENTREGAS", 0), errors="coerce").sum()
    total_coletas = pd.to_numeric(df_viag.get("COLETAS", 0), errors="coerce").sum()
    total_pedidos = pd.to_numeric(df_viag.get("PEDIDOS", 0), errors="coerce").sum()
    taxa_sucesso = (total_entregas / total_pedidos * 100) if total_pedidos > 0 else 0

    v1, v2, v3, v4 = st.columns(4)
    with v1:
        st.markdown(_bi_card("Viagens", str(int(total_viagens)), "#8b5cf6", "🚛"), unsafe_allow_html=True)
    with v2:
        st.markdown(_bi_card("KM Rodado", f"{km_total:,.0f} km".replace(",", "."), "#33CCFF", "🛣️"), unsafe_allow_html=True)
    with v3:
        color_taxa = "#10b981" if taxa_sucesso >= 85 else "#f59e0b" if taxa_sucesso >= 70 else "#ef4444"
        st.markdown(_bi_card("Taxa Sucesso", f"{taxa_sucesso:.1f}%", color_taxa, "📈"), unsafe_allow_html=True)
    with v4:
        st.markdown(_bi_card("Coletas", str(int(total_coletas)), "#ec4899", "📥"), unsafe_allow_html=True)

    # Gráfico de KM por motorista
    if "MOTORISTA" in df_viag.columns:
        km_motor = (
            df_viag.groupby("MOTORISTA")
            .agg(KM=("KM_RODADO", lambda x: pd.to_numeric(x, errors="coerce").sum()),
                 VIAGENS=("DATA", "count"))
            .reset_index()
            .sort_values("KM", ascending=False)
        )
        if not km_motor.empty and km_motor["KM"].sum() > 0:
            fig = px.bar(
                km_motor, x="MOTORISTA", y="KM",
                color="VIAGENS",
                color_continuous_scale=[[0, "#1e1b4b"], [0.5, "#8b5cf6"], [1, "#c4b5fd"]],
                title="KM Rodado por Motorista",
                labels={"KM": "KM Total", "MOTORISTA": "", "VIAGENS": "Viagens"},
                text=km_motor["KM"].apply(lambda v: f"{v:,.0f} km".replace(",", ".")),
            )
            fig.update_traces(textposition="outside", marker_cornerradius=6)
            fig.update_layout(**PLOTLY_LAYOUT, height=350, coloraxis_showscale=False)
            st.plotly_chart(fig, use_container_width=True)

    st.markdown("---")

# ══════════════════════════════════════════════════════════════
# 6. VEÍCULOS + DISTRIBUIÇÃO GEOGRÁFICA
# ══════════════════════════════════════════════════════════════
st.markdown(
    '<div style="color: #10b981; font-size: 13px; font-weight: 700; '
    'text-transform: uppercase; letter-spacing: 1px; margin-bottom: 12px;">'
    '🚛 Veículos & Distribuição</div>',
    unsafe_allow_html=True,
)

c_veic, c_geo = st.columns(2)

with c_veic:
    sv = get_summary_by_vehicle(df_f)
    sv_fleet = sv[sv["VEICULO"].isin(FLEET_PLATES)]
    if not sv_fleet.empty:
        fig = px.bar(
            sv_fleet, x="VEICULO", y="TOTAL_NFS",
            color="VEICULO", color_discrete_sequence=COLORS,
            title="Entregas por Veículo (Frota)",
            labels={"TOTAL_NFS": "Total NFs", "VEICULO": "Placa"},
            text="TOTAL_NFS",
        )
        fig.update_traces(
            textposition="outside",
            hovertemplate="<b>Placa:</b> %{x}<br><b>NFs:</b> %{y}<extra></extra>",
            marker_cornerradius=6,
        )
        fig.update_layout(**PLOTLY_LAYOUT, height=400, showlegend=False)
        st.plotly_chart(fig, use_container_width=True)

with c_geo:
    uf = get_deliveries_by_uf(df_f)
    if not uf.empty and len(uf) > 1:
        fig = px.treemap(
            uf, path=["UF"], values="TOTAL_NFS",
            color="VALOR_TOTAL",
            color_continuous_scale=[[0, "#064e3b"], [0.5, "#10b981"], [1, "#6ee7b7"]],
            title="Distribuição Geográfica (UF)",
        )
        fig.update_traces(
            hovertemplate=(
                "<b>UF:</b> %{label}<br>"
                "<b>NFs:</b> %{value}<br>"
                "<b>Valor Total:</b> R$ %{color:,.2f}<extra></extra>"
            ),
        )
        fig.update_layout(**PLOTLY_LAYOUT, height=400, coloraxis_showscale=False)
        st.plotly_chart(fig, use_container_width=True)

# ── Donut de operações ──
if not df_f.empty:
    ops_df = df_f["OPERACAO"].value_counts().reset_index()
    ops_df.columns = ["OPERACAO", "TOTAL"]
    fig = px.pie(
        ops_df, names="OPERACAO", values="TOTAL",
        color_discrete_sequence=COLORS,
        title="Distribuição por Tipo de Operação",
        hole=0.55,
    )
    fig.update_traces(
        hovertemplate="<b>%{label}</b><br>Total: %{value} (%{percent})<extra></extra>",
        textposition="inside",
        textinfo="percent+label",
    )
    fig.update_layout(**PLOTLY_LAYOUT, height=380, showlegend=False)
    st.plotly_chart(fig, use_container_width=True)

# ══════════════════════════════════════════════════════════════
# 7. CONTROLE NOTAS — STATUS DOS PEDIDOS
# ══════════════════════════════════════════════════════════════
if not df_pedidos.empty and "STATUS" in df_pedidos.columns:
    st.markdown("---")
    st.markdown(
        '<div style="color: #10b981; font-size: 13px; font-weight: 700; '
        'text-transform: uppercase; letter-spacing: 1px; margin-bottom: 12px;">'
        '📋 Status dos Pedidos (Controle Notas)</div>',
        unsafe_allow_html=True,
    )

    status_counts = df_pedidos["STATUS"].value_counts()
    total_pedidos_cn = len(df_pedidos)

    # KPIs de status
    st_cols = st.columns(min(len(status_counts), 5))
    for i, (status, count) in enumerate(status_counts.items()):
        col_idx = i % len(st_cols)
        color = STATUS_COLORS.get(status, "#10b981")
        with st_cols[col_idx]:
            st.markdown(
                _bi_card(status, str(count), color, ""),
                unsafe_allow_html=True,
            )

    # Gráfico de barras de status
    status_df = status_counts.reset_index()
    status_df.columns = ["STATUS", "TOTAL"]
    colors_mapped = [STATUS_COLORS.get(s, "#10b981") for s in status_df["STATUS"]]

    fig = px.bar(
        status_df, x="STATUS", y="TOTAL",
        title="Distribuição de Pedidos por Status",
        text="TOTAL",
        color="STATUS",
        color_discrete_map=STATUS_COLORS,
    )
    fig.update_traces(textposition="outside", marker_cornerradius=6)
    fig.update_layout(**PLOTLY_LAYOUT, height=350, showlegend=False)
    st.plotly_chart(fig, use_container_width=True)

    # Por empresa
    if "EMPRESA" in df_pedidos.columns:
        emp_status = df_pedidos.groupby(["EMPRESA", "STATUS"]).size().reset_index(name="TOTAL")
        if not emp_status.empty:
            fig = px.bar(
                emp_status, x="EMPRESA", y="TOTAL", color="STATUS",
                color_discrete_map=STATUS_COLORS,
                title="Pedidos por Empresa × Status",
                barmode="stack",
            )
            fig.update_traces(marker_cornerradius=4)
            fig.update_layout(**PLOTLY_LAYOUT, height=350)
            st.plotly_chart(fig, use_container_width=True)
