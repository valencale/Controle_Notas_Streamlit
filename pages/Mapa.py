"""
Mapa.py — Geolocalização de entregas e pedidos.

Seções:
    1. KPIs: Top bairros, total de localidades, volume por UF
    2. Mapa de calor (density) das entregas por bairro/UF
    3. Mapa de pontos dos pedidos ativos (geocoded)
    4. Rankings e tabelas de apoio

Fontes de dados:
    - RELATÓRIO DE ENTREGAS (via delivery_reader) → UF + Bairro (sem lat/lon)
    - CONTROLE NOTAS (.xlsm) → Endereço (geocoded para lat/lon)
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import folium
from folium.plugins import HeatMap
from streamlit_folium import st_folium
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from modules.ui_components import inject_custom_css, render_header
from modules.excel_handler import read_principal
from modules.geocoder import geocode_address
from config import STATUS_COLORS, STATUS_ICONS, MAP_CENTER, MAP_ZOOM

# Try to import delivery_reader
try:
    from modules.delivery_reader import read_deliveries_report
    HAS_DELIVERY_REPORT = True
except ImportError:
    HAS_DELIVERY_REPORT = False

# ══════════════════════════════════════════════════════════════
# PAGE CONFIG
# ══════════════════════════════════════════════════════════════
inject_custom_css()

# Plotly dark theme
PLOTLY_LAYOUT = dict(
    template="plotly_dark",
    paper_bgcolor="rgba(0,0,0,0)",
    plot_bgcolor="rgba(0,0,0,0)",
    font=dict(family="Inter, sans-serif", color="#cbd5e1", size=13),
    margin=dict(l=20, r=20, t=50, b=20),
    title_font=dict(size=17, color="#f8fafc"),
    hoverlabel=dict(
        bgcolor="rgba(15,23,42,0.95)",
        font_size=13,
        font_family="Inter, sans-serif",
        bordercolor="rgba(16,185,129,0.4)",
    ),
)


def _hex_to_rgb(hex_color):
    h = hex_color.lstrip("#")
    if len(h) == 6:
        return f"{int(h[0:2], 16)}, {int(h[2:4], 16)}, {int(h[4:6], 16)}"
    return "16, 185, 129"


def _geo_card(label, value, color, icon=""):
    """KPI card for geo dashboard."""
    return (
        '<div style="'
        f'background: linear-gradient(135deg, rgba({_hex_to_rgb(color)}, 0.12), rgba({_hex_to_rgb(color)}, 0.04));'
        f'border: 1px solid rgba({_hex_to_rgb(color)}, 0.2);'
        f'border-left: 4px solid {color};'
        'border-radius: 14px;'
        'padding: 16px 20px;'
        'box-shadow: 0 4px 16px rgba(0,0,0,0.15);'
        '">'
        f'<div style="color: {color}; font-size: 11px; font-weight: 700;'
        f' text-transform: uppercase; letter-spacing: 0.8px; margin-bottom: 6px;">'
        f'{icon} {label}</div>'
        f'<div style="color: #f8fafc; font-size: 26px; font-weight: 700; line-height: 1.2;">'
        f'{value}</div>'
        '</div>'
    )


# ══════════════════════════════════════════════════════════════
# DATA LOADING
# ══════════════════════════════════════════════════════════════
@st.cache_data(show_spinner="Carregando relatório de entregas...", ttl=300)
def load_deliveries():
    if not HAS_DELIVERY_REPORT:
        return pd.DataFrame()
    try:
        return read_deliveries_report()
    except FileNotFoundError:
        return pd.DataFrame()


# Coordenadas aproximadas das capitais/regiões brasileiras por UF
UF_COORDS = {
    "SP": (-23.55, -46.63), "RJ": (-22.91, -43.17), "MG": (-19.92, -43.94),
    "ES": (-20.32, -40.34), "PR": (-25.43, -49.27), "SC": (-27.59, -48.55),
    "RS": (-30.03, -51.23), "BA": (-12.97, -38.51), "PE": (-8.05, -34.87),
    "CE": (-3.72, -38.52), "PA": (-1.46, -48.50), "AM": (-3.12, -60.02),
    "GO": (-16.69, -49.25), "DF": (-15.78, -47.93), "MT": (-15.60, -56.10),
    "MS": (-20.44, -54.65), "MA": (-2.53, -44.28), "PI": (-5.09, -42.80),
    "RN": (-5.79, -35.21), "PB": (-7.12, -34.86), "AL": (-9.67, -35.74),
    "SE": (-10.91, -37.07), "TO": (-10.18, -48.33), "RO": (-8.76, -63.90),
    "AC": (-9.97, -67.81), "AP": (0.03, -51.07), "RR": (2.82, -60.67),
}

# Bairros mais comuns de SP com coordenadas aproximadas
SP_BAIRROS = {
    "OSASCO": (-23.5325, -46.7917), "BARUERI": (-23.5114, -46.8761),
    "ALPHAVILLE": (-23.4946, -46.8494), "CARAPICUÍBA": (-23.5225, -46.8358),
    "COTIA": (-23.6037, -46.9192), "ITAPEVI": (-23.5488, -46.9342),
    "JANDIRA": (-23.5278, -46.9025), "SANTANA PARNAÍBA": (-23.4439, -46.9175),
    "PINHEIROS": (-23.5631, -46.6917), "MOEMA": (-23.5997, -46.6650),
    "ITAIM BIBI": (-23.5855, -46.6753), "VILA MARIANA": (-23.5897, -46.6361),
    "LAPA": (-23.5194, -46.7017), "PERDIZES": (-23.5350, -46.6850),
    "CENTRO": (-23.5505, -46.6340), "BARRA FUNDA": (-23.5250, -46.6650),
    "VILA OLIMPIA": (-23.5961, -46.6864), "BROOKLIN": (-23.6097, -46.6797),
    "SANTO AMARO": (-23.6533, -46.7092), "CAMPO BELO": (-23.6200, -46.6650),
    "TATUAPE": (-23.5381, -46.5764), "PENHA": (-23.5194, -46.5444),
    "SANTANA": (-23.5036, -46.6256), "TUCURUVI": (-23.4797, -46.6031),
    "VILA PRUDENTE": (-23.5803, -46.5819), "IPIRANGA": (-23.5872, -46.6025),
    "JABAQUARA": (-23.6339, -46.6425), "SAÚDE": (-23.6167, -46.6358),
    "BUTANTÃ": (-23.5722, -46.7367), "MORUMBI": (-23.6097, -46.7208),
    "GUARULHOS": (-23.4628, -46.5333), "SÃO BERNARDO": (-23.6939, -46.5650),
    "SANTO ANDRE": (-23.6739, -46.5433), "DIADEMA": (-23.6861, -46.6228),
    "MAUÁ": (-23.6678, -46.4614), "CAMPINAS": (-22.9064, -47.0616),
    "SOROCABA": (-23.5015, -47.4526), "RIBEIRÃO PRETO": (-21.1775, -47.8103),
    "SÃO JOSÉ DOS CAMPOS": (-23.1896, -45.8841),
}


# ══════════════════════════════════════════════════════════════
# HEADER
# ══════════════════════════════════════════════════════════════
render_header("Mapa de Entregas", "Geolocalização e análise de volume por região")

# ══════════════════════════════════════════════════════════════
# LOAD DATA
# ══════════════════════════════════════════════════════════════
df_del = load_deliveries()

if df_del.empty:
    st.warning("Relatório de entregas não encontrado. Atualize o cache na sidebar de Administração.")
    st.stop()

# ══════════════════════════════════════════════════════════════
# FILTERS (sidebar)
# ══════════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown("### 🔧 Filtros Mapa")

    # Período
    if "DATA" in df_del.columns and df_del["DATA"].notna().any():
        min_date = df_del["DATA"].min().date()
        max_date = df_del["DATA"].max().date()
        periodo = st.date_input(
            "📅 Período", value=(min_date, max_date),
            min_value=min_date, max_value=max_date, key="map_periodo",
        )
    else:
        periodo = None

    # UF
    ufs = sorted(df_del["UF"].unique().tolist())
    sel_uf = st.multiselect("🗺️ Estado (UF)", ufs, default=ufs, key="map_uf")

    # Remetente
    remetentes = sorted(df_del["REMETENTE"].unique().tolist())
    sel_rem = st.multiselect("📦 Remetente", remetentes, default=remetentes, key="map_rem")

# Apply filters
mask = df_del["UF"].isin(sel_uf) & df_del["REMETENTE"].isin(sel_rem)
if periodo and len(periodo) == 2:
    mask = mask & (df_del["DATA"].dt.date >= periodo[0]) & (df_del["DATA"].dt.date <= periodo[1])
df_f = df_del[mask].copy()

if df_f.empty:
    st.info("Nenhum dado encontrado para os filtros selecionados.")
    st.stop()

# ══════════════════════════════════════════════════════════════
# 1. KPIs GEOGRÁFICOS
# ══════════════════════════════════════════════════════════════
st.markdown(
    '<div style="color: #10b981; font-size: 13px; font-weight: 700; '
    'text-transform: uppercase; letter-spacing: 1px; margin-bottom: 12px;">'
    '📍 Indicadores Geográficos</div>',
    unsafe_allow_html=True,
)

# Agrupar por bairro
bairro_agg = (
    df_f.groupby("BAIRRO")
    .agg(NFs=("NOTA_FISCAL", "count"), PESO=("PESO", "sum"), VALOR=("VALOR_NOTA", "sum"))
    .reset_index()
    .sort_values("NFs", ascending=False)
)
bairro_agg = bairro_agg[bairro_agg["BAIRRO"].str.strip() != ""]

# Agrupar por UF
uf_agg = (
    df_f.groupby("UF")
    .agg(NFs=("NOTA_FISCAL", "count"), PESO=("PESO", "sum"), VALOR=("VALOR_NOTA", "sum"))
    .reset_index()
    .sort_values("NFs", ascending=False)
)

total_localidades = len(bairro_agg)
total_ufs = len(uf_agg)
top_bairro = bairro_agg.iloc[0]["BAIRRO"] if not bairro_agg.empty else "—"
top_bairro_nfs = int(bairro_agg.iloc[0]["NFs"]) if not bairro_agg.empty else 0

g1, g2, g3, g4 = st.columns(4)
with g1:
    st.markdown(_geo_card("Bairros Visitados", str(total_localidades), "#10b981", "🏘️"), unsafe_allow_html=True)
with g2:
    st.markdown(_geo_card("Estados Atendidos", str(total_ufs), "#33CCFF", "🗺️"), unsafe_allow_html=True)
with g3:
    st.markdown(_geo_card("Top Bairro", top_bairro, "#f59e0b", "🏆"), unsafe_allow_html=True)
with g4:
    st.markdown(_geo_card(f"NFs em {top_bairro}", str(top_bairro_nfs), "#6366f1", "📄"), unsafe_allow_html=True)

# Top 5 bairros
if len(bairro_agg) >= 5:
    st.markdown(
        '<div style="color: #94a3b8; font-size: 11px; font-weight: 600; '
        'text-transform: uppercase; letter-spacing: 0.8px; margin: 12px 0 6px 0;">'
        '🥇 Top 5 Bairros com Mais Entregas</div>',
        unsafe_allow_html=True,
    )
    top5 = bairro_agg.head(5)
    t_cols = st.columns(5)
    top5_colors = ["#10b981", "#06d6a0", "#33CCFF", "#6366f1", "#f59e0b"]
    for i, (_, row) in enumerate(top5.iterrows()):
        with t_cols[i]:
            st.markdown(
                _geo_card(row["BAIRRO"], f'{int(row["NFs"])} NFs', top5_colors[i], f"#{i+1}"),
                unsafe_allow_html=True,
            )

st.markdown("---")

# ══════════════════════════════════════════════════════════════
# 2. MAPA DE CALOR + PONTOS DE ENTREGA
# ══════════════════════════════════════════════════════════════
st.markdown(
    '<div style="color: #10b981; font-size: 13px; font-weight: 700; '
    'text-transform: uppercase; letter-spacing: 1px; margin-bottom: 12px;">'
    '🔥 Mapa de Calor — Volume de Entregas</div>',
    unsafe_allow_html=True,
)

# Gerar coordenadas para heatmap usando bairros conhecidos e UFs
heat_data = []
marker_data = []

for _, row in bairro_agg.iterrows():
    bairro_upper = str(row["BAIRRO"]).strip().upper()
    nfs = int(row["NFs"])

    # Tentar match com bairros conhecidos de SP
    coords = SP_BAIRROS.get(bairro_upper)
    if not coords:
        # Tentar variações parciais
        for key, val in SP_BAIRROS.items():
            if key in bairro_upper or bairro_upper in key:
                coords = val
                break

    if coords:
        heat_data.append([coords[0], coords[1], nfs])
        marker_data.append({
            "BAIRRO": row["BAIRRO"], "NFs": nfs,
            "PESO": row["PESO"], "VALOR": row["VALOR"],
            "LAT": coords[0], "LON": coords[1],
        })

# Adicionar pontos por UF para bairros sem match
for _, row in uf_agg.iterrows():
    uf_upper = str(row["UF"]).strip().upper()
    coords = UF_COORDS.get(uf_upper)
    if coords:
        # Calcular quantas NFs desse UF já foram mapeadas por bairro
        mapped_nfs = sum(
            m["NFs"] for m in marker_data
            # Sem como cruzar UF→bairro aqui, adiciona UF como fallback
        )
        heat_data.append([coords[0], coords[1], int(row["NFs"]) * 0.3])

# Criar mapa Folium (tiles=None para esconder nome no LayerControl)
m = folium.Map(
    location=MAP_CENTER,
    zoom_start=10,
    tiles=None,
    control_scale=True,
)
folium.TileLayer("CartoDB dark_matter", name="Mapa Base").add_to(m)

# Camada de calor (dentro de FeatureGroup nomeado)
if heat_data:
    fg_heat = folium.FeatureGroup(name="🔥 Mapa de Calor")
    HeatMap(
        heat_data,
        radius=25,
        blur=18,
        max_zoom=13,
        gradient={0.2: '#10b981', 0.4: '#06d6a0', 0.6: '#f59e0b', 0.8: '#ef4444', 1.0: '#dc2626'},
    ).add_to(fg_heat)
    fg_heat.add_to(m)

# Marcadores dos bairros com volume
if marker_data:
    fg_bairros = folium.FeatureGroup(name="📍 Bairros com Entregas")
    for md in marker_data:
        popup_html = f"""
        <div style="font-family: 'Inter', sans-serif; min-width: 180px;">
            <h4 style="margin: 0 0 8px 0; color: #10b981;">{md['BAIRRO']}</h4>
            <p style="margin: 3px 0;"><strong>NFs:</strong> {md['NFs']}</p>
            <p style="margin: 3px 0;"><strong>Peso:</strong> {md['PESO']:,.0f} kg</p>
            <p style="margin: 3px 0;"><strong>Valor:</strong> R$ {md['VALOR']:,.2f}</p>
        </div>
        """
        # Tamanho do círculo proporcional ao volume
        radius = max(6, min(25, md["NFs"] / 5))
        folium.CircleMarker(
            location=[md["LAT"], md["LON"]],
            radius=radius,
            color="#10b981",
            fill=True,
            fill_color="#10b981",
            fill_opacity=0.6,
            popup=folium.Popup(popup_html, max_width=250),
            tooltip=f"{md['BAIRRO']}: {md['NFs']} NFs",
        ).add_to(fg_bairros)
    fg_bairros.add_to(m)

# ── Pedidos ativos do Controle Notas (geocoded) ──
try:
    df_pedidos = read_principal()
    df_with_addr = df_pedidos[
        df_pedidos["ENDERECO"].notna() & (df_pedidos["ENDERECO"].astype(str).str.strip() != "")
    ].copy()

    if not df_with_addr.empty:
        if "geocode_cache" not in st.session_state:
            st.session_state.geocode_cache = {}

        FOLIUM_COLORS = {
            "SEPARACAO": "orange", "PARCIAL": "beige", "AUSENTE": "red",
            "CONCLUIDO": "green", "AGUARDANDO NF": "pink",
            "SEM MATERIAL": "red", "ENVIAR DATA": "purple",
            "EM ROTA": "blue", "ENTREGUE": "darkgreen",
        }

        fg_pedidos = folium.FeatureGroup(name="📋 Pedidos Ativos")
        geocoded_count = 0

        for _, row in df_with_addr.iterrows():
            addr = str(row["ENDERECO"]).strip()
            if addr in st.session_state.geocode_cache:
                coords = st.session_state.geocode_cache[addr]
            else:
                coords = geocode_address(addr)
                st.session_state.geocode_cache[addr] = coords

            if coords:
                geocoded_count += 1
                status = row.get("STATUS", "SEPARACAO")
                color = FOLIUM_COLORS.get(status, "blue")
                icon_char = STATUS_ICONS.get(status, "📋")

                data_str = ""
                try:
                    data_str = pd.Timestamp(row["DATA"]).strftime("%d/%m/%Y")
                except Exception:
                    data_str = str(row.get("DATA", ""))

                popup_html = f"""
                <div style="font-family: 'Inter', sans-serif; min-width: 200px;">
                    <h4 style="margin: 0 0 8px 0; color: #1e293b;">Pedido #{row['PEDIDO']}</h4>
                    <p style="margin: 3px 0;"><strong>Cliente:</strong> {row['CLIENTE']}</p>
                    <p style="margin: 3px 0;"><strong>Status:</strong> {icon_char} {status}</p>
                    <p style="margin: 3px 0;"><strong>Data:</strong> {data_str}</p>
                    <p style="margin: 3px 0; font-size: 11px; color: #64748b;">📍 {addr}</p>
                </div>
                """
                folium.Marker(
                    location=[coords[0], coords[1]],
                    popup=folium.Popup(popup_html, max_width=300),
                    tooltip=f"#{row['PEDIDO']} — {row['CLIENTE']}",
                    icon=folium.Icon(color=color, icon="home", prefix="glyphicon"),
                ).add_to(fg_pedidos)

        if geocoded_count > 0:
            fg_pedidos.add_to(m)

except Exception:
    pass  # Controle Notas indisponível, sem problema

# Layer control
folium.LayerControl(collapsed=False).add_to(m)

# Render
st_folium(m, width=None, height=600, use_container_width=True)

st.markdown("---")

# ══════════════════════════════════════════════════════════════
# 3. GRÁFICOS DE DISTRIBUIÇÃO
# ══════════════════════════════════════════════════════════════
st.markdown(
    '<div style="color: #10b981; font-size: 13px; font-weight: 700; '
    'text-transform: uppercase; letter-spacing: 1px; margin-bottom: 12px;">'
    '📊 Distribuição Geográfica</div>',
    unsafe_allow_html=True,
)

c_uf, c_bairro = st.columns(2)

# ── Volume por UF ──
with c_uf:
    if not uf_agg.empty:
        fig = px.bar(
            uf_agg.head(10), x="UF", y="NFs",
            color="VALOR",
            color_continuous_scale=[[0, "#064e3b"], [0.5, "#10b981"], [1, "#6ee7b7"]],
            title="Volume de Entregas por UF",
            labels={"NFs": "Total NFs", "UF": "Estado", "VALOR": "Valor (R$)"},
            text="NFs",
        )
        fig.update_traces(textposition="outside", marker_cornerradius=6)
        fig.update_layout(**PLOTLY_LAYOUT, height=400, coloraxis_showscale=False)
        st.plotly_chart(fig, use_container_width=True)

# ── Top Bairros ──
with c_bairro:
    if not bairro_agg.empty:
        top_bairros = bairro_agg.head(12)
        fig = px.bar(
            top_bairros, x="NFs", y="BAIRRO", orientation="h",
            color="PESO",
            color_continuous_scale=[[0, "#312e81"], [0.5, "#6366f1"], [1, "#a5b4fc"]],
            title="Top 12 Bairros por Entregas",
            labels={"NFs": "Total NFs", "BAIRRO": "", "PESO": "Peso (kg)"},
            text="NFs",
        )
        fig.update_traces(textposition="outside")
        fig.update_layout(**PLOTLY_LAYOUT, height=400, coloraxis_showscale=False)
        fig.update_yaxes(categoryorder="total ascending")
        st.plotly_chart(fig, use_container_width=True)

# ── Treemap UF → Bairro ──
if not df_f.empty:
    tree_data = (
        df_f[df_f["BAIRRO"].str.strip() != ""]
        .groupby(["UF", "BAIRRO"])
        .agg(NFs=("NOTA_FISCAL", "count"), VALOR=("VALOR_NOTA", "sum"))
        .reset_index()
        .sort_values("NFs", ascending=False)
        .head(50)
    )
    if not tree_data.empty and len(tree_data) > 1:
        fig = px.treemap(
            tree_data, path=["UF", "BAIRRO"], values="NFs",
            color="VALOR",
            color_continuous_scale=[[0, "#064e3b"], [0.5, "#10b981"], [1, "#6ee7b7"]],
            title="Mapa Hierárquico: UF → Bairro",
        )
        fig.update_traces(
            hovertemplate="<b>%{label}</b><br>NFs: %{value}<br>Valor: R$ %{color:,.2f}<extra></extra>",
        )
        fig.update_layout(**PLOTLY_LAYOUT, height=450, coloraxis_showscale=False)
        st.plotly_chart(fig, use_container_width=True)

# ── Tabela completa ──
with st.expander("📋 Tabela de entregas por bairro"):
    if not bairro_agg.empty:
        display = bairro_agg.copy()
        display.columns = ["Bairro", "NFs", "Peso (kg)", "Valor (R$)"]
        st.dataframe(
            display.style.format({
                "Peso (kg)": "{:,.0f}",
                "Valor (R$)": "R$ {:,.2f}",
            }).background_gradient(subset=["NFs"], cmap="Greens"),
            use_container_width=True,
            height=400,
        )
