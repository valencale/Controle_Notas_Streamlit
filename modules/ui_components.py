"""
ui_components.py — Componentes de UI reutilizáveis para a aplicação Streamlit.

Contém:
- Barra de KPIs com st.metric
- CSS customizado para tema logístico dark
- Helpers de feedback (success/error)
- Card rendering
"""

import streamlit as st
import pandas as pd

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from config import STATUS_COLORS, STATUS_OPTIONS


def inject_custom_css():
    """Injeta CSS customizado a partir do arquivo assets/styles.css para visual premium."""
    css_path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "assets", "styles.css")
    try:
        with open(css_path, "r", encoding="utf-8") as f:
            css_content = f.read()
        st.markdown(f"<style>{css_content}</style>", unsafe_allow_html=True)
    except Exception as e:
        st.warning(f"Não foi possível carregar os estilos customizados: {e}")



def render_header(title: str, subtitle: str = ""):
    """Renderiza o header principal da página."""
    st.markdown(f"""
    <div class="main-header animate-in">
        <div>
            <h1>{title}</h1>
            <div class="subtitle">{subtitle}</div>
        </div>
    </div>
    """, unsafe_allow_html=True)


def render_kpi_bar(df: pd.DataFrame, historico_counts: dict = None):
    """
    Renderiza a barra de KPIs com contagem por status, usando cards coloridos.

    Args:
        df: DataFrame da aba principal (Dados).
        historico_counts: Dict opcional {status: count} com contagens do Histórico.
                          Usado para somar EM ROTA e ENTREGUE que já foram movidos.
    """
    if historico_counts is None:
        historico_counts = {}

    total = len(df)
    status_counts = df["STATUS"].value_counts().to_dict() if not df.empty else {}

    # Soma contagens do Histórico para EM ROTA e ENTREGUE
    for status_key in ("EM ROTA", "ENTREGUE"):
        hist_val = historico_counts.get(status_key, 0)
        if hist_val > 0:
            status_counts[status_key] = status_counts.get(status_key, 0) + hist_val
            total += hist_val

    items = [("TOTAL PEDIDOS", total, "#6366f1")]
    for status in STATUS_OPTIONS:
        count = status_counts.get(status, 0)
        color = STATUS_COLORS.get(status, "#6366f1")
        items.append((status, count, color))

    col_count = 3
    for i in range(0, len(items), col_count):
        cols = st.columns(col_count)
        for j in range(col_count):
            if i + j < len(items):
                label, val, color = items[i + j]
                with cols[j]:
                    card_html = (
                        '<div style="'
                        f'background: linear-gradient(135deg, {color}22, {color}11);'
                        f'border: 1px solid {color}44;'
                        f'border-left: 4px solid {color};'
                        'border-radius: 12px;'
                        'padding: 16px 20px;'
                        '">'
                        f'<div style="color: {color}; font-size: 12px; font-weight: 700;'
                        ' text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 4px;">'
                        f'{label}</div>'
                        f'<div style="color: var(--gb-text-primary); font-size: 32px; font-weight: 700;">'
                        f'{val}</div>'
                        '</div>'
                    )
                    st.markdown(card_html, unsafe_allow_html=True)


def render_kpi_bar_compact(df: pd.DataFrame, historico_counts: dict = None):
    """
    Renderiza KPIs compactos em 2 linhas horizontais.

    Linha 1: TOTAL + SEPARACAO + PARCIAL + AUSENTE + CONCLUIDO
    Linha 2: AGUARDANDO NF + SEM MATERIAL + EM ROTA + ENTREGUE

    Args:
        df: DataFrame da aba principal (Dados).
        historico_counts: Dict opcional com contagens do Histórico.
    """
    from config import STATUS_ICONS

    if historico_counts is None:
        historico_counts = {}

    total = len(df)
    status_counts = df["STATUS"].value_counts().to_dict() if not df.empty else {}

    # Soma contagens do Histórico para EM ROTA e ENTREGUE
    for status_key in ("EM ROTA", "ENTREGUE"):
        hist_val = historico_counts.get(status_key, 0)
        if hist_val > 0:
            status_counts[status_key] = status_counts.get(status_key, 0) + hist_val
            total += hist_val

    # Monta itens: (label, valor, cor, ícone)
    items = [("TOTAL", total, "#6366f1", "📦")]
    for status in STATUS_OPTIONS:
        count = status_counts.get(status, 0)
        color = STATUS_COLORS.get(status, "#6366f1")
        icon = STATUS_ICONS.get(status, "")
        # Label abreviado para caber
        short = {
            "SEPARACAO": "SEPARAÇÃO",
            "PARCIAL": "PARCIAL",
            "AUSENTE": "AUSENTE",
            "CONCLUIDO": "CONCLUÍDO",
            "AGUARDANDO NF": "AGUARDANDO NF",
            "SEM MATERIAL": "SEM MATERIAL",
            "ENVIAR DATA": "ENVIAR DATA",
            "EM ROTA": "EM ROTA",
            "ENTREGUE": "ENTREGUE",
        }.get(status, status)
        items.append((short, count, color, icon))

    # Renderiza em 2 linhas (5 + 5)
    row1 = items[:5]   # TOTAL + 4 primeiros status
    row2 = items[5:]   # 5 últimos status

    for row_items in [row1, row2]:
        cols = st.columns(len(row_items))
        for idx, (label, val, color, icon) in enumerate(row_items):
            with cols[idx]:
                badge_html = (
                    f'<div style="'
                    f'background: {color}15;'
                    f'border: 1px solid {color}33;'
                    f'border-left: 3px solid {color};'
                    f'border-radius: 8px;'
                    f'padding: 6px 10px;'
                    f'display: flex; align-items: center; gap: 8px;'
                    f'">'
                    f'<span style="font-size: 14px;">{icon}</span>'
                    f'<span style="color: {color}; font-size: 11px; font-weight: 700;'
                    f' text-transform: uppercase; letter-spacing: 0.3px;">{label}</span>'
                    f'<span style="color: var(--gb-text-primary); font-size: 18px; font-weight: 700;'
                    f' margin-left: auto;">{val}</span>'
                    f'</div>'
                )
                st.markdown(badge_html, unsafe_allow_html=True)


def render_status_badge(status: str) -> str:
    """Retorna HTML de um badge de status colorido."""
    color = STATUS_COLORS.get(status, "#6366f1")
    return f'<span class="status-badge" style="background: {color}22; color: {color}; border: 1px solid {color}44;">{status}</span>'


def render_pedido_card_html(row: dict) -> str:
    """Retorna HTML de um card de pedido (apenas visual)."""

    def _safe(val, default=""):
        """Sanitize value: treat None, nan, 'nan', 'None' as empty."""
        if pd.isna(val):
            return default
        s = str(val).strip()
        if s.lower() in ("nan", "none", "nat", "<na>"):
            return default
        return s

    status_html = render_status_badge(_safe(row.get("STATUS"), ""))
    data_str = ""
    if row.get("DATA"):
        try:
            ts = pd.Timestamp(row["DATA"])
            if pd.notna(ts):
                data_str = ts.strftime("%d/%m/%Y")
        except Exception:
            data_str = _safe(row.get("DATA"))

    import re
    obs_val = _safe(row.get("OBS"))
    # Remove eventuais tags HTML que tenham sido salvas acidentalmente no campo
    if obs_val:
        obs_val = re.sub(r'<[^>]+>', '', obs_val)
    obs_html = f'<div class="pedido-obs">{obs_val}</div>' if obs_val else ""

    endereco_val = _safe(row.get("ENDERECO"))
    endereco_html = f'<div style="font-size: 13px; color: #94a3b8; margin-top: 6px; display: flex; align-items: center; gap: 6px;"><svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M20 10c0 6-8 12-8 12s-8-6-8-12a8 8 0 0 1 16 0Z"/><circle cx="12" cy="10" r="3"/></svg>{endereco_val}</div>' if endereco_val else ""

    pedido_val = _safe(row.get("PEDIDO"), "N/A")
    cliente_val = _safe(row.get("CLIENTE"), "N/A")
    empresa_val = _safe(row.get("EMPRESA"))

    return f"""
<div class="pedido-card animate-in">
    <div style="display: flex; justify-content: space-between; align-items: flex-start;">
        <div style="display: flex; align-items: baseline; gap: 12px;">
            <span class="pedido-number">#{pedido_val}</span>
            <span class="pedido-empresa">{empresa_val}</span>
            <span style="color: #64748b; font-size: 13px;">{data_str}</span>
        </div>
        {status_html}
    </div>
    <div class="pedido-cliente">{cliente_val}</div>
    {endereco_html}
    {obs_html}
</div>
    """

