"""
2_📦_Historico.py — Página de visualização do Histórico e Estorno.

Funcionalidades:
- Visualização dos pedidos arquivados
- Busca e filtro por data
- Estorno (retornar pedido à esteira ativa)
"""

import streamlit as st
import pandas as pd
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from modules.ui_components import inject_custom_css, render_header, render_status_badge
from modules.excel_handler import read_historico, restore_from_historico
from modules.auth import is_adm

# ══════════════════════════════════════════════════════════════
# PAGE CONFIG
# ══════════════════════════════════════════════════════════════
st.set_page_config(page_title="Histórico — Gestão Logística", page_icon="📦", layout="wide")
inject_custom_css()

# ══════════════════════════════════════════════════════════════
# HEADER
# ══════════════════════════════════════════════════════════════
render_header("Histórico de Pedidos", "Pedidos arquivados (concluídos)")

st.markdown("---")

# ══════════════════════════════════════════════════════════════
# LOAD DATA
# ══════════════════════════════════════════════════════════════
try:
    df = read_historico()
except Exception as e:
    st.error(f"❌ Erro ao carregar histórico: {e}")
    st.stop()

if df.empty:
    st.markdown("""
    <div style="text-align: center; padding: 60px 0;">
        <div style="font-size: 72px; margin-bottom: 16px;">📭</div>
        <h3 style="color: #94a3b8; font-weight: 600;">Histórico vazio</h3>
        <p style="color: #64748b; font-size: 14px;">
            Nenhum pedido foi arquivado ainda.<br>
            Use o botão "📦 Arquivar Concluídos" na Esteira para mover pedidos concluídos.
        </p>
    </div>
    """, unsafe_allow_html=True)
    st.stop()

# ══════════════════════════════════════════════════════════════
# KPIs
# ══════════════════════════════════════════════════════════════
col1, col2, col3 = st.columns(3)
with col1:
    st.metric("📦 Total Arquivados", len(df))
with col2:
    empresas = df["EMPRESA"].nunique() if not df.empty else 0
    st.metric("🏢 Empresas", empresas)
with col3:
    clientes = df["CLIENTE"].nunique() if not df.empty else 0
    st.metric("👥 Clientes Únicos", clientes)

st.markdown("---")

# ══════════════════════════════════════════════════════════════
# SEARCH + FILTER
# ══════════════════════════════════════════════════════════════
col_search, col_filter = st.columns([3, 1])

with col_search:
    search_query = st.text_input(
        "🔍 Buscar no Histórico",
        placeholder="Pesquisar por Pedido ou Cliente...",
        label_visibility="collapsed",
        key="search_historico"
    )

with col_filter:
    empresa_filter = st.selectbox(
        "Filtrar por Empresa",
        ["Todas"] + sorted(df["EMPRESA"].dropna().unique().tolist()),
        label_visibility="collapsed",
    )

# Apply filters
filtered_df = df.copy()

if search_query:
    query_upper = search_query.upper().strip()
    mask = (
        filtered_df["PEDIDO"].astype(str).str.contains(query_upper, case=False, na=False) |
        filtered_df["CLIENTE"].astype(str).str.contains(query_upper, case=False, na=False)
    )
    filtered_df = filtered_df[mask]

if empresa_filter != "Todas":
    filtered_df = filtered_df[filtered_df["EMPRESA"] == empresa_filter]

# ══════════════════════════════════════════════════════════════
# DATA TABLE
# ══════════════════════════════════════════════════════════════
st.markdown(f"<p style='color: #64748b; font-size: 13px;'>Exibindo {len(filtered_df)} registro(s)</p>", unsafe_allow_html=True)

# Format the dataframe for display
display_df = filtered_df.copy()
if "DATA" in display_df.columns:
    display_df["DATA"] = pd.to_datetime(display_df["DATA"], errors="coerce").dt.strftime("%d/%m/%Y")

st.dataframe(
    display_df,
    width="stretch",
    hide_index=True,
    column_config={
        "PEDIDO": st.column_config.TextColumn("📋 Pedido", width="small"),
        "CLIENTE": st.column_config.TextColumn("👤 Cliente", width="medium"),
        "DATA": st.column_config.TextColumn("📅 Data", width="small"),
        "EMPRESA": st.column_config.TextColumn("🏢 Empresa", width="small"),
        "STATUS": st.column_config.TextColumn("🔖 Status", width="small"),
        "OBS": st.column_config.TextColumn("📝 Obs", width="medium"),
        "ENDERECO": st.column_config.TextColumn("📍 Endereço", width="medium"),
        "NF": st.column_config.TextColumn("📄 NF", width="small"),
    },
)

# ══════════════════════════════════════════════════════════════
# ESTORNO (Restore to Esteira)
# ══════════════════════════════════════════════════════════════
st.markdown("---")
st.markdown("### ↩️ Estorno — Retornar Pedido à Esteira")

st.markdown("""
<p style="color: #64748b; font-size: 13px;">
    Selecione um pedido para devolvê-lo à esteira ativa. O status será resetado para <strong>SEPARACAO</strong>.
</p>
""", unsafe_allow_html=True)

pedido_options = filtered_df["PEDIDO"].tolist()

if pedido_options:
    selected_pedido = st.selectbox(
        "Selecionar Pedido para Estorno",
        pedido_options,
        format_func=lambda x: f"#{x} — {filtered_df[filtered_df['PEDIDO'] == x]['CLIENTE'].values[0] if len(filtered_df[filtered_df['PEDIDO'] == x]) > 0 else 'N/A'}",
    )

    if st.button("↩️ Retornar à Esteira", type="primary", disabled=not is_adm()):
        try:
            restore_from_historico(str(selected_pedido))
            st.success(f"✅ Pedido #{selected_pedido} retornado à esteira com status SEPARACAO!")
            st.rerun()
        except Exception as e:
            st.error(f"❌ Erro no estorno: {e}")
else:
    st.info("Nenhum pedido disponível para estorno.")
