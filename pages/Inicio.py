"""
app.py — Entry point da aplicação Streamlit de Gestão Logística de Armazém.

Este é o ponto de entrada principal. Configura o layout, injeta CSS global
e renderiza a sidebar com navegação. As páginas são auto-descobertas
pelo Streamlit a partir da pasta /pages.
"""

import streamlit as st
import sys
import os

# Add project root to path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from modules.ui_components import inject_custom_css, render_header

# ══════════════════════════════════════════════════════════════
# PAGE CONFIG — Deve ser o primeiro comando Streamlit
# ══════════════════════════════════════════════════════════════
st.set_page_config(
    page_title="Gestão Logística — Armazém",
    page_icon="🏭",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ══════════════════════════════════════════════════════════════
# CSS GLOBAL
# ══════════════════════════════════════════════════════════════
inject_custom_css()

# ══════════════════════════════════════════════════════════════
# AUTENTICAÇÃO
# ══════════════════════════════════════════════════════════════
from modules.auth import init_auth, login_screen

init_auth()
if not login_screen():
    st.stop()

# ══════════════════════════════════════════════════════════════
# SIDEBAR
# ══════════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown("""
    <div style="text-align: center; padding: 20px 0;">
        <svg xmlns="http://www.w3.org/2000/svg" width="48" height="48" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" style="color: #818cf8; margin-bottom: 8px;">
            <path d="M2 22h20"/><path d="M17 2v20"/><path d="M7 2v20"/><path d="M12 22v-6"/><path d="M7 12h10"/><path d="M7 6h10"/>
        </svg>
        <h1 style="margin: 0; font-size: 22px;">Controle Notas</h1>
        <p style="color: var(--gb-text-muted, #64748b); font-size: 13px; margin-top: 4px;">
            Sistema de Gestão Logística
        </p>
    </div>
    """, unsafe_allow_html=True)

    st.divider()

    st.divider()

    st.markdown("""
    <div style="padding: 8px 0; color: var(--gb-text-muted, #475569); font-size: 11px;">
        <strong>Backend:</strong> CONTROLE NOTAS.xlsm<br>
        <strong>Motor:</strong> openpyxl + Streamlit<br>
        <strong>Versão:</strong> 1.0.0
    </div>
    """, unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════
# HOME PAGE
# ══════════════════════════════════════════════════════════════
render_header("Gestão Logística de Armazém", "Painel operacional integrado com Excel")

# ── Load data ──
from modules.excel_handler import read_principal
from config import STATUS_OPTIONS, STATUS_COLORS

try:
    df = read_principal()
except Exception as e:
    st.error(f"Erro ao carregar dados: {e}", icon=":material/error:")
    st.info("Verifique se o ficheiro CONTROLE NOTAS.xlsm está acessível.")
    st.stop()




# ══════════════════════════════════════════════════════════════
# SECTION 1: ESTEIRA OVERVIEW — Subtle Colored KPIs
# ══════════════════════════════════════════════════════════════
st.markdown("""
<div style="margin-bottom: 8px;">
    <span style="color: var(--gb-text-muted, #94a3b8); font-size: 12px; font-weight: 600; text-transform: uppercase; letter-spacing: 1px;">
        📊 Visão Geral — Esteira de Pedidos
    </span>
</div>
""", unsafe_allow_html=True)

total = len(df)
status_counts = df["STATUS"].value_counts().to_dict() if not df.empty else {}


def _subtle_card(label, value, color, opacity="0.08"):
    """Card KPI com cor sutil — não agressivo, legível."""
    return (
        '<div style="'
        f'background: rgba({_hex_to_rgb(color)}, {opacity});'
        f'border: 1px solid rgba({_hex_to_rgb(color)}, 0.2);'
        f'border-left: 3px solid {color};'
        'border-radius: 10px;'
        'padding: 14px 18px;'
        '">'
        f'<div style="color: {color}; font-size: 11px; font-weight: 600;'
        ' text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 2px;'
        f' opacity: 0.85;">{label}</div>'
        f'<div style="color: var(--gb-text-primary, #e2e8f0); font-size: 28px; font-weight: 700;">'
        f'{value}</div>'
        '</div>'
    )


def _hex_to_rgb(hex_color):
    """Convert #RRGGBB to 'R, G, B' string for rgba()."""
    h = hex_color.lstrip("#")
    return f"{int(h[0:2], 16)}, {int(h[2:4], 16)}, {int(h[4:6], 16)}"


# Row 1: Total + main statuses
items = [("TOTAL", total, "#6366f1")]
for s in STATUS_OPTIONS:
    items.append((s, status_counts.get(s, 0), STATUS_COLORS.get(s, "#6366f1")))

# Dynamic columns based on number of items (max 4 per row)
col_count = min(len(items), 4)
for i in range(0, len(items), col_count):
    cols = st.columns(col_count)
    for j in range(col_count):
        if i + j < len(items):
            label, val, color = items[i + j]
            with cols[j]:
                st.markdown(_subtle_card(label, val, color), unsafe_allow_html=True)

st.markdown("---")

# ══════════════════════════════════════════════════════════════
# SECTION 2: APP DESCRIPTION
# ══════════════════════════════════════════════════════════════
st.markdown("""
<div style="
    background: linear-gradient(135deg, rgba(99, 102, 241, 0.06), rgba(99, 102, 241, 0.02));
    border: 1px solid rgba(99, 102, 241, 0.12);
    border-radius: 14px;
    padding: 24px 28px;
    margin-bottom: 24px;
">
    <div style="display: flex; align-items: center; gap: 12px; margin-bottom: 12px;">
        <svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none"
             stroke="#818cf8" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
            <circle cx="12" cy="12" r="10"/><path d="M12 16v-4"/><path d="M12 8h.01"/>
        </svg>
        <span style="color: var(--gb-text-primary, #e2e8f0); font-size: 16px; font-weight: 600;">Sobre o Sistema</span>
    </div>
    <p style="color: var(--gb-text-muted, #94a3b8); font-size: 14px; line-height: 1.7; margin: 0;">
        O <strong style="color: var(--gb-accent, #c7d2fe);">Controle Notas</strong> é um painel de gestão logística
        que centraliza o acompanhamento de pedidos, separação de mercadorias, ingestão de notas fiscais,
        roteirização de entregas e análise de performance — tudo integrado com sua planilha Excel operacional.
        Utilize os módulos abaixo para acessar cada funcionalidade.
    </p>
</div>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════
# SECTION 3: FEATURE NAVIGATION CARDS
# ══════════════════════════════════════════════════════════════
st.markdown("""
<div style="margin-bottom: 12px;">
    <span style="color: var(--gb-text-muted, #94a3b8); font-size: 12px; font-weight: 600; text-transform: uppercase; letter-spacing: 1px;">
        🚀 Módulos Disponíveis
    </span>
</div>
""", unsafe_allow_html=True)

# Feature cards data
features = [
    {
        "icon": "🏭",
        "title": "Esteira de Pedidos",
        "desc": "Gerencie pedidos ativos com CRUD completo, status coloridos e busca em tempo real.",
        "color": "#818cf8",
        "page": "pages/Esteira.py",
    },
    {
        "icon": "📁",
        "title": "Histórico",
        "desc": "Consulte pedidos arquivados, filtre por período e exporte relatórios.",
        "color": "#f59e0b",
        "page": "pages/Historico.py",
    },
    {
        "icon": "📄",
        "title": "Ingestão PDF",
        "desc": "Importe notas fiscais de arquivos PDF com extração automática de dados.",
        "color": "#10b981",
        "page": "pages/Ingestao_PDF.py",
    },
    {
        "icon": "🗺️",
        "title": "Mapa de Entregas",
        "desc": "Visualize rotas de entrega com geolocalização e otimização de percurso.",
        "color": "#3b82f6",
        "page": "pages/Mapa.py",
    },
    {
        "icon": "📋",
        "title": "Checklist Separação",
        "desc": "Controle multi-picking de mercadorias com conferência de estoque.",
        "color": "#ec4899",
        "page": "pages/Checklist_Separacao.py",
    },
    {
        "icon": "📝",
        "title": "Extração Notas",
        "desc": "Leitura e extração de dados de DANFE/XML em lote.",
        "color": "#f97316",
        "page": "pages/Extracao_Notas.py",
    },
    {
        "icon": "📊",
        "title": "BI & Analytics",
        "desc": "Dashboards de entregas, KPIs de performance logística e análise de frotas.",
        "color": "#8b5cf6",
        "page": "pages/BI_Analytics.py",
    },
    {
        "icon": "🚚",
        "title": "Expedição Diária",
        "desc": "Planejamento de saída, checklist imprimível, registro de viagens, auditoria e conferência.",
        "color": "#22d3ee",
        "page": "pages/Expedicao.py",
    },
    {
        "icon": "📆",
        "title": "Histórico Expedições",
        "desc": "Consulte o histórico completo de rotas e expedições fechadas.",
        "color": "#a855f7",
        "page": "pages/Historico_Expedicoes.py",
    },
]


def _feature_card(feat):
    """Glassmorphism-style feature card."""
    c = feat["color"]
    return (
        f'<div style="'
        f'background: linear-gradient(135deg, rgba({_hex_to_rgb(c)}, 0.06), rgba({_hex_to_rgb(c)}, 0.02));'
        f'border: 1px solid rgba({_hex_to_rgb(c)}, 0.15);'
        'border-radius: 14px;'
        'padding: 20px;'
        'height: 100%;'
        'transition: all 0.2s ease;'
        '">'
        f'<div style="font-size: 28px; margin-bottom: 8px;">{feat["icon"]}</div>'
        f'<div style="color: var(--gb-text-primary, #f1f5f9); font-size: 15px; font-weight: 600; margin-bottom: 6px;">'
        f'{feat["title"]}</div>'
        f'<div style="color: var(--gb-text-muted, #94a3b8); font-size: 13px; line-height: 1.5;">'
        f'{feat["desc"]}</div>'
        '</div>'
    )


# Render feature cards in 3-column grid with buttons
for i in range(0, len(features), 3):
    cols = st.columns(3)
    for j in range(3):
        if i + j < len(features):
            feat = features[i + j]
            with cols[j]:
                st.markdown(_feature_card(feat), unsafe_allow_html=True)
                st.page_link(
                    feat["page"],
                    label=f"Abrir {feat['title']}",
                    icon=":material/arrow_forward:",
                    use_container_width=True,
                )
    if i + 3 < len(features):
        st.markdown("<div style='margin-bottom: 16px;'></div>", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════
# FOOTER
# ══════════════════════════════════════════════════════════════
st.markdown("---")
st.markdown("""
<div style="text-align: center; padding: 16px 0; color: var(--gb-text-muted, #475569); font-size: 12px;">
    <span style="opacity: 0.7;">Controle Notas v1.1 • Powered by Streamlit + openpyxl</span>
</div>
""", unsafe_allow_html=True)
