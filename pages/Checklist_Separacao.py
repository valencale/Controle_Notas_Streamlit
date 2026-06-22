"""
5_📋_Checklist_Separacao.py — Checklist de Separação Multi-Picking.

Cruza estoque (PDF/Word) com pedidos (PDF) e gera painéis de:
- Batch Picking (visão por produto/SKU)
- Discrete Picking (visão por pedido/caixa)
"""

import streamlit as st
import pandas as pd
import io
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from modules.ui_components import inject_custom_css, render_header
from modules.checklist_engine import (
    extrair_estoque_pdf,
    extrair_estoque_docx,
    extrair_pedidos_pdf,
    cruzar_dados,
)
from modules.auth import is_adm

# ══════════════════════════════════════════════════════════════
# PAGE CONFIG
# ══════════════════════════════════════════════════════════════
st.set_page_config(page_title="Checklist de Separação", page_icon="📋", layout="wide")
inject_custom_css()

# CSS adicional para formatação condicional
st.markdown("""
<style>
    .kpi-row { display: flex; gap: 12px; margin-bottom: 20px; }
    .kpi-box {
        flex: 1; padding: 18px 20px; border-radius: 14px;
        background: linear-gradient(135deg, #1e293b, #1a1f2e);
        border: 1px solid rgba(99,102,241,0.15);
        text-align: center;
    }
    .kpi-box .kpi-val { font-size: 32px; font-weight: 700; color: #f8fafc; }
    .kpi-box .kpi-lab { font-size: 12px; color: #94a3b8; text-transform: uppercase;
        letter-spacing: 0.5px; margin-top: 4px; }
    .kpi-green { border-left: 4px solid #22c55e; }
    .kpi-amber { border-left: 4px solid #f59e0b; }
    .kpi-red   { border-left: 4px solid #ef4444; }
    .kpi-blue  { border-left: 4px solid #6366f1; }
    .expander-completo summary { color: #22c55e !important; }
    .expander-parcial  summary { color: #ef4444 !important; }
</style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════
# HEADER
# ══════════════════════════════════════════════════════════════
render_header("Checklist de Separação", "Cruzamento Estoque × Pedidos — Batch & Discrete Picking")
st.markdown("---")

# ══════════════════════════════════════════════════════════════
# INGESTÃO DE ARQUIVOS
# ══════════════════════════════════════════════════════════════
col_est, col_ped = st.columns(2)
with col_est:
    st.markdown("### 📦 Posição de Estoque")
    arquivo_estoque = st.file_uploader(
        "Upload do inventário (.pdf ou .docx)",
        type=["pdf", "docx"],
        key="uploader_estoque",
        help="PDF CIGAM ou Word com tabela de estoque",
        disabled=not is_adm(),
    )
with col_ped:
    st.markdown("### 📄 Pedidos / Mapas de Separação")
    arquivos_pedidos = st.file_uploader(
        "Upload dos pedidos (.pdf)",
        type=["pdf"],
        accept_multiple_files=True,
        key="uploader_pedidos",
        help="PDFs Onfinity — Mapa de Separação por Pedido",
        disabled=not is_adm(),
    )

st.markdown("---")

# ══════════════════════════════════════════════════════════════
# BOTÃO DE PROCESSAMENTO
# ══════════════════════════════════════════════════════════════
btn_disabled = not (arquivo_estoque and arquivos_pedidos)
btn = st.button(
    "🚀 Gerar Mapa de Separação",
    use_container_width=True,
    type="primary",
    disabled=btn_disabled or not is_adm(),
)

if btn_disabled and not btn:
    st.markdown("""
    <div style="text-align:center; padding:50px 0;">
        <div style="font-size:64px; margin-bottom:12px;">📋</div>
        <h3 style="color:#94a3b8;">Anexe os arquivos acima para começar</h3>
        <p style="color:#64748b; font-size:14px;">
            1️⃣ Posição de Estoque (PDF ou Word) &nbsp;→&nbsp;
            2️⃣ Pedidos (um ou mais PDFs)
        </p>
    </div>
    """, unsafe_allow_html=True)
    st.stop()

if not btn:
    # Se já processou antes, mostra resultados do session_state
    if "checklist_result" not in st.session_state:
        st.stop()

# ══════════════════════════════════════════════════════════════
# PROCESSAMENTO
# ══════════════════════════════════════════════════════════════
if btn:
    with st.status("🔍 Processando arquivos...", expanded=True) as status:
        # 1. Estoque
        st.write("📦 Lendo posição de estoque...")
        nome_est = arquivo_estoque.name.lower()
        try:
            if nome_est.endswith(".docx"):
                df_estoque = extrair_estoque_docx(arquivo_estoque)
            else:
                df_estoque = extrair_estoque_pdf(arquivo_estoque)
        except Exception as e:
            st.error(f"❌ Erro ao ler estoque: {e}")
            st.stop()

        if df_estoque.empty:
            status.update(label="❌ Estoque vazio", state="error")
            st.error("Não foi possível extrair dados do arquivo de estoque.")
            st.stop()
        st.write(f"✅ {len(df_estoque)} SKUs de estoque carregados")

        # 2. Pedidos
        st.write("📄 Lendo pedidos...")
        try:
            df_pedidos = extrair_pedidos_pdf(arquivos_pedidos)
        except Exception as e:
            st.error(f"❌ Erro ao ler pedidos: {e}")
            st.stop()

        if df_pedidos.empty:
            status.update(label="❌ Pedidos vazios", state="error")
            st.error("Não foi possível extrair itens dos pedidos.")
            st.stop()

        n_pedidos = df_pedidos["Pedido"].nunique()
        n_itens = len(df_pedidos)
        st.write(f"✅ {n_pedidos} pedidos com {n_itens} itens extraídos")

        # Verificar se há pedidos "Desconhecido" residuais (não deveria mais ocorrer)
        n_desconhecido = len(df_pedidos[df_pedidos["Pedido"] == "Desconhecido"])
        if n_desconhecido > 0:
            st.warning(
                f"⚠️ {n_desconhecido} item(ns) sem pedido identificado. "
                f"Verifique se o PDF possui páginas com formato inesperado."
            )

        # 3. Cruzamento
        st.write("🔀 Cruzando estoque com pedidos...")
        resultado = cruzar_dados(df_estoque, df_pedidos)
        st.session_state.checklist_result = resultado

        status.update(label=f"✅ Processado: {n_pedidos} pedidos × {len(df_estoque)} SKUs", state="complete")

    st.toast(f"✅ {n_pedidos} pedidos processados com sucesso!", icon="📋")

# ══════════════════════════════════════════════════════════════
# RESULTADOS
# ══════════════════════════════════════════════════════════════
if "checklist_result" not in st.session_state:
    st.stop()

res = st.session_state.checklist_result
stats = res["stats"]
batch_df = res["batch_df"]

# KPIs
st.markdown(f"""
<div class="kpi-row">
    <div class="kpi-box kpi-blue">
        <div class="kpi-val">{stats['total_pedidos']}</div>
        <div class="kpi-lab">Pedidos Lidos</div>
    </div>
    <div class="kpi-box kpi-green">
        <div class="kpi-val">{stats['pedidos_completos']}</div>
        <div class="kpi-lab">Completos</div>
    </div>
    <div class="kpi-box kpi-amber">
        <div class="kpi-val">{stats['pedidos_parciais']}</div>
        <div class="kpi-lab">Parciais</div>
    </div>
    <div class="kpi-box kpi-red">
        <div class="kpi-val">{stats['pedidos_sem_estoque']}</div>
        <div class="kpi-lab">Sem Estoque</div>
    </div>
    <div class="kpi-box kpi-blue">
        <div class="kpi-val">{stats['total_skus']}</div>
        <div class="kpi-lab">SKUs Únicos</div>
    </div>
</div>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════
# ABAS: BATCH vs DISCRETE
# ══════════════════════════════════════════════════════════════
aba_batch, aba_discrete = st.tabs(["🛒 Separação Agrupada (Batch Picking)", "📦 Separação por Pedido (Discrete Picking)"])

# ── ABA 1: BATCH PICKING ─────────────────────────────────────
with aba_batch:
    st.subheader("Itens Consolidados — Visão por Produto")
    st.caption("Busque todos estes itens no galpão e depois distribua nas caixas dos pedidos.")

    # Formatação condicional
    def _highlight_falta(row):
        if row["Falta"] > 0:
            return ["background-color: rgba(239,68,68,0.15); color: #fca5a5;"] * len(row)
        return [""] * len(row)

    df_display = batch_df[["Local", "Código", "Descrição", "Estoque_Atual", "Qtd_Total", "Falta", "Status", "Num_Pedidos", "Pedidos_Destino"]].copy()
    df_display.columns = ["Local", "Código", "Descrição", "Estoque Atual", "Total Necessário", "Falta", "Status", "Nº Pedidos", "Distribuir Para"]

    styled = df_display.style.apply(_highlight_falta, axis=1)
    st.dataframe(styled, use_container_width=True, hide_index=True, height=500)

    # Exportação
    st.markdown("#### 📥 Exportar")
    col_csv, col_xlsx = st.columns(2)
    with col_csv:
        csv_data = df_display.to_csv(index=False).encode("utf-8-sig")
        st.download_button(
            "📄 Baixar CSV",
            csv_data,
            "batch_picking.csv",
            "text/csv",
            use_container_width=True,
        )
    with col_xlsx:
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            df_display.to_excel(writer, index=False, sheet_name="Batch Picking")
        st.download_button(
            "📊 Baixar Excel",
            buf.getvalue(),
            "batch_picking.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

# ── ABA 2: DISCRETE PICKING ──────────────────────────────────
with aba_discrete:
    st.subheader("Conferência por Pedido — Visão por Caixa")
    st.caption("Cada pedido é uma caixa. Verifique item por item antes de despachar.")

    # --- Pedidos Completos ---
    completos = res["pedidos_completos"]
    if completos:
        st.markdown(f"### 🟢 Pedidos Completos ({len(completos)})")
        for info in completos:
            with st.expander(f"✅ Pedido {info['pedido']} — {info['cliente']} ({info['total_itens']} itens)"):
                st.dataframe(
                    info["itens_df"][["Local", "Código", "Descrição", "Qtd_Pedida"]],
                    use_container_width=True, hide_index=True,
                )

    # --- Pedidos Parciais ---
    parciais = res["pedidos_parciais"]
    if parciais:
        st.markdown(f"### 🟡 Pedidos Parciais ({len(parciais)})")
        for info in parciais:
            pct = int(info["itens_ok"] / info["total_itens"] * 100) if info["total_itens"] else 0
            with st.expander(f"⚠️ Pedido {info['pedido']} — {info['cliente']} ({pct}% disponível)"):
                st.dataframe(
                    info["itens_df"].style.apply(
                        lambda r: ["background-color: rgba(239,68,68,0.15);"] * len(r) if r["Falta"] > 0 else [""] * len(r),
                        axis=1
                    ),
                    use_container_width=True, hide_index=True,
                )

    # --- Pedidos Sem Estoque ---
    sem_estoque = res["pedidos_sem_estoque"]
    if sem_estoque:
        st.markdown(f"### 🔴 Pedidos Sem Estoque ({len(sem_estoque)})")
        for info in sem_estoque:
            with st.expander(f"❌ Pedido {info['pedido']} — {info['cliente']} (0% disponível)"):
                st.dataframe(
                    info["itens_df"][["Local", "Código", "Descrição", "Qtd_Pedida", "Estoque_Atual"]],
                    use_container_width=True, hide_index=True,
                )

    # Exportação Discrete
    st.markdown("---")
    st.markdown("#### 📥 Exportar Todos os Pedidos")
    buf2 = io.BytesIO()
    with pd.ExcelWriter(buf2, engine="openpyxl") as writer:
        # Aba resumo
        resumo_rows = []
        for info in completos:
            resumo_rows.append({"Pedido": info["pedido"], "Cliente": info["cliente"], "Status": "COMPLETO", "Itens": info["total_itens"]})
        for info in parciais:
            pct = int(info["itens_ok"] / info["total_itens"] * 100) if info["total_itens"] else 0
            resumo_rows.append({"Pedido": info["pedido"], "Cliente": info["cliente"], "Status": f"PARCIAL ({pct}%)", "Itens": info["total_itens"]})
        for info in sem_estoque:
            resumo_rows.append({"Pedido": info["pedido"], "Cliente": info["cliente"], "Status": "SEM ESTOQUE", "Itens": info["total_itens"]})
        if resumo_rows:
            pd.DataFrame(resumo_rows).to_excel(writer, index=False, sheet_name="Resumo")

        # Aba com todos os itens
        master = res["master_df"].copy()
        master["Falta"] = (master["Qtd_Pedida"] - master["Estoque_Atual"]).clip(lower=0)
        master.to_excel(writer, index=False, sheet_name="Todos os Itens")

    st.download_button(
        "📊 Baixar Excel Completo (Todos os Pedidos)",
        buf2.getvalue(),
        "discrete_picking.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )