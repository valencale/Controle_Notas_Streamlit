"""
4_🗺️_Mapa.py — Página de roteirização geoespacial.

Funcionalidades:
- Mapa interativo centrado em São Paulo/Osasco
- Marcadores coloridos por status para pedidos com endereço
- Popup com informações do pedido em cada marcador
- Controle de camadas por status
"""

import streamlit as st
import pandas as pd
import folium
from streamlit_folium import st_folium
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from modules.ui_components import inject_custom_css, render_header
from modules.excel_handler import read_principal
from modules.geocoder import geocode_address, get_default_center
from config import STATUS_COLORS, STATUS_ICONS, MAP_CENTER, MAP_ZOOM

# ══════════════════════════════════════════════════════════════
# PAGE CONFIG
# ══════════════════════════════════════════════════════════════
st.set_page_config(page_title="Mapa — Gestão Logística", page_icon="🗺️", layout="wide")
inject_custom_css()

# ══════════════════════════════════════════════════════════════
# HEADER
# ══════════════════════════════════════════════════════════════
render_header("Mapa de Roteirização", "Visualização geoespacial dos pedidos ativos")

st.markdown("---")

# ══════════════════════════════════════════════════════════════
# LOAD DATA
# ══════════════════════════════════════════════════════════════
try:
    df = read_principal()
except Exception as e:
    st.error(f"❌ Erro ao carregar dados: {e}")
    st.stop()

# Filter only rows with address
df_with_addr = df[df["ENDERECO"].notna() & (df["ENDERECO"].astype(str).str.strip() != "")].copy()

if df_with_addr.empty:
    st.markdown("""
    <div style="text-align: center; padding: 60px 0;">
        <div style="font-size: 72px; margin-bottom: 16px;">🗺️</div>
        <h3 style="color: #94a3b8; font-weight: 600;">Nenhum pedido com endereço</h3>
        <p style="color: #64748b; font-size: 14px;">
            Importe PDFs com endereço na página de <strong>Ingestão PDF</strong><br>
            ou adicione endereços manualmente na <strong>Esteira</strong>.
        </p>
    </div>
    """, unsafe_allow_html=True)

    # Still show an empty map centered on SP
    m = folium.Map(location=MAP_CENTER, zoom_start=MAP_ZOOM, tiles="CartoDB dark_matter")
    st_folium(m, width=None, height=500, use_container_width=True)
    st.stop()

# ══════════════════════════════════════════════════════════════
# KPIs
# ══════════════════════════════════════════════════════════════
col1, col2, col3 = st.columns(3)
with col1:
    st.metric("📍 Pedidos com Endereço", len(df_with_addr))
with col2:
    st.metric("📦 Total Pedidos Ativos", len(df))
with col3:
    pct = int(len(df_with_addr) / len(df) * 100) if len(df) > 0 else 0
    st.metric("📊 Cobertura", f"{pct}%")

st.markdown("---")

# ══════════════════════════════════════════════════════════════
# GEOCODING
# ══════════════════════════════════════════════════════════════

# Initialize geocode cache in session state
if "geocode_cache" not in st.session_state:
    st.session_state.geocode_cache = {}

# Folium color mapping for status
FOLIUM_COLORS = {
    "SEPARACAO": "orange",
    "PARCIAL": "beige",
    "AUSENTE": "red",
    "CONCLUIDO": "green",
    "AGUARDANDO NF": "pink",
}

# Geocode addresses
with st.spinner("🌍 Geocodificando endereços..."):
    geocoded_rows = []
    for idx, row in df_with_addr.iterrows():
        addr = str(row["ENDERECO"]).strip()

        # Check session state cache first
        if addr in st.session_state.geocode_cache:
            coords = st.session_state.geocode_cache[addr]
        else:
            coords = geocode_address(addr)
            st.session_state.geocode_cache[addr] = coords

        if coords:
            geocoded_rows.append({
                **row.to_dict(),
                "LAT": coords[0],
                "LON": coords[1],
            })

if not geocoded_rows:
    st.warning("⚠️ Nenhum endereço pôde ser geocodificado. Verifique os endereços dos pedidos.")
    m = folium.Map(location=MAP_CENTER, zoom_start=MAP_ZOOM, tiles="CartoDB dark_matter")
    st_folium(m, width=None, height=500, use_container_width=True)
    st.stop()

geocoded_df = pd.DataFrame(geocoded_rows)

# ══════════════════════════════════════════════════════════════
# MAP
# ══════════════════════════════════════════════════════════════
st.markdown(f"### 🗺️ {len(geocoded_df)} pedido(s) no mapa")

# Create map
m = folium.Map(
    location=MAP_CENTER,
    zoom_start=MAP_ZOOM,
    tiles="CartoDB dark_matter",
)

# Add markers grouped by status (as feature groups for layer control)
status_groups = {}
for status in geocoded_df["STATUS"].unique():
    fg = folium.FeatureGroup(name=f"{STATUS_ICONS.get(status, '📋')} {status}")
    status_groups[status] = fg

for idx, row in geocoded_df.iterrows():
    status = row.get("STATUS", "SEPARACAO")
    color = FOLIUM_COLORS.get(status, "blue")
    icon_char = STATUS_ICONS.get(status, "📋")

    # Format date
    data_str = ""
    try:
        data_str = pd.Timestamp(row["DATA"]).strftime("%d/%m/%Y")
    except Exception:
        data_str = str(row.get("DATA", ""))

    # Popup HTML
    popup_html = f"""
    <div style="font-family: 'Inter', sans-serif; min-width: 200px;">
        <h4 style="margin: 0 0 8px 0; color: #1e293b;">Pedido #{row['PEDIDO']}</h4>
        <p style="margin: 4px 0;"><strong>Cliente:</strong> {row['CLIENTE']}</p>
        <p style="margin: 4px 0;"><strong>Empresa:</strong> {row.get('EMPRESA', 'N/A')}</p>
        <p style="margin: 4px 0;"><strong>Status:</strong> {status}</p>
        <p style="margin: 4px 0;"><strong>Data:</strong> {data_str}</p>
        <p style="margin: 4px 0; font-size: 11px; color: #64748b;">📍 {row['ENDERECO']}</p>
    </div>
    """

    marker = folium.Marker(
        location=[row["LAT"], row["LON"]],
        popup=folium.Popup(popup_html, max_width=300),
        tooltip=f"#{row['PEDIDO']} — {row['CLIENTE']}",
        icon=folium.Icon(color=color, icon="info-sign"),
    )

    fg = status_groups.get(status)
    if fg:
        marker.add_to(fg)

# Add all feature groups to map
for fg in status_groups.values():
    fg.add_to(m)

# Layer control
folium.LayerControl(collapsed=False).add_to(m)

# Render map
map_data = st_folium(m, width=None, height=600, use_container_width=True)

# ══════════════════════════════════════════════════════════════
# ADDRESS TABLE
# ══════════════════════════════════════════════════════════════
st.markdown("---")

with st.expander("📋 Lista de Endereços Geocodificados", expanded=False):
    display_cols = ["PEDIDO", "CLIENTE", "EMPRESA", "STATUS", "ENDERECO", "LAT", "LON"]
    available_cols = [c for c in display_cols if c in geocoded_df.columns]
    st.dataframe(
        geocoded_df[available_cols],
        width="stretch",
        hide_index=True,
    )
