"""
1_🏭_Esteira.py — Página principal: Tabela editável inline + Busca + Gestão de Status.

Funcionalidades:
- KPIs compactos (2 linhas)
- Busca por Pedido OU Cliente (em tempo real)
- Tabela editável inline (st.data_editor) — como Excel
- Inserção de novos pedidos
- Exclusão em lote (checkbox + botão)
- Arquivamento de pedidos EM ROTA
"""

import streamlit as st
import pandas as pd
from datetime import datetime
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from modules.ui_components import inject_custom_css, render_header, render_kpi_bar_compact
from modules.excel_handler import (
    read_principal,
    insert_pedido,
    update_rows_batch,
    delete_pedidos_batch,
    archive_completed,
    archive_dispatched,
    count_status_in_historico,
)
from st_keyup import st_keyup
from config import STATUS_OPTIONS, EMPRESA_OPTIONS, STATUS_COLORS
from modules.auth import is_adm

# ══════════════════════════════════════════════════════════════
# PAGE CONFIG
# ══════════════════════════════════════════════════════════════
st.set_page_config(page_title="Esteira — Gestão Logística", page_icon="🏭", layout="wide")
inject_custom_css()

# ══════════════════════════════════════════════════════════════
# STATUS EMOJI MAPPING (coloração visual no data_editor)
# ══════════════════════════════════════════════════════════════
STATUS_EMOJI = {
    "SEPARACAO": "🟡",
    "PARCIAL": "🟠",
    "AUSENTE": "🔴",
    "CONCLUIDO": "🟢",
    "AGUARDANDO NF": "💗",
    "SEM MATERIAL": "⛔",
    "ENVIAR DATA": "🟣",
    "EM ROTA": "🔵",
    "ENTREGUE": "✅",
}

# Opções com emoji para o SelectboxColumn
STATUS_DISPLAY_OPTIONS = [f"{STATUS_EMOJI.get(s, '')} {s}" for s in STATUS_OPTIONS]

# Raw → Display
def _status_to_display(raw):
    raw_clean = str(raw).strip().upper()
    emoji = STATUS_EMOJI.get(raw_clean, "")
    return f"{emoji} {raw_clean}" if emoji else raw_clean

# Display → Raw (strip emoji)
def _display_to_status(display):
    s = str(display).strip()
    for raw_name in STATUS_EMOJI:
        if raw_name in s.upper():
            return raw_name
    # Fallback: remove any leading emoji chars
    return s.lstrip("🟡🟠🔴🟢💗⛔🟣🔵✅ ").strip()

# ══════════════════════════════════════════════════════════════
# STATE INITIALIZATION
# ══════════════════════════════════════════════════════════════
if "refresh_trigger" not in st.session_state:
    st.session_state.refresh_trigger = 0

if "show_new_form" not in st.session_state:
    st.session_state.show_new_form = False


def trigger_refresh():
    st.session_state.refresh_trigger += 1


# ══════════════════════════════════════════════════════════════
# LOAD DATA (cacheado em session_state — só relê do disco quando necessário)
# ══════════════════════════════════════════════════════════════
try:
    need_reload = (
        "df_pedidos" not in st.session_state
        or st.session_state.refresh_trigger > 0
    )
    if need_reload:
        st.session_state.df_pedidos = read_principal()
        st.session_state.historico_counts = count_status_in_historico()
        st.session_state.refresh_trigger = 0
    df = st.session_state.df_pedidos
except Exception as e:
    st.error(f"Erro ao carregar Excel: {e}")
    st.stop()

# ══════════════════════════════════════════════════════════════
# HEADER + KPIs COMPACTOS
# ══════════════════════════════════════════════════════════════
render_header("Esteira de Pedidos", "Gestão em tempo real dos pedidos ativos")

historico_counts = st.session_state.get("historico_counts", {})
render_kpi_bar_compact(df, historico_counts=historico_counts)

st.markdown("---")

# ══════════════════════════════════════════════════════════════
# SEARCH + ACTIONS BAR
# ══════════════════════════════════════════════════════════════
col_search, col_actions = st.columns([3, 2], vertical_alignment="bottom")

with col_search:
    search_query = st_keyup(
        "Buscar",
        placeholder="Pesquisar por número do Pedido ou nome do Cliente...",
        label_visibility="collapsed",
        key="search_esteira"
    )

with col_actions:
    btn_col1, btn_col2 = st.columns(2)
    with btn_col1:
        if st.button("Novo Pedido", icon=":material/add:", width="stretch", type="primary", disabled=not is_adm()):
            st.session_state.show_new_form = not st.session_state.show_new_form
    with btn_col2:
        em_rota_count = len(df[df["STATUS"] == "EM ROTA"]) if not df.empty else 0
        if st.button(
            f"Arquivar EM ROTA ({em_rota_count})",
            icon=":material/local_shipping:",
            width="stretch",
            disabled=em_rota_count == 0 or not is_adm(),
        ):
            st.session_state.show_archive_confirm = True

# ══════════════════════════════════════════════════════════════
# ARCHIVE CONFIRMATION
# ══════════════════════════════════════════════════════════════
if st.session_state.get("show_archive_confirm", False):
    st.warning(f"Tem certeza que deseja mover **{em_rota_count}** pedido(s) EM ROTA para o Histórico?")
    c1, c2, c3 = st.columns([1, 1, 4])
    with c1:
        if st.button("Sim, arquivar", type="primary", icon=":material/check:", disabled=not is_adm()):
            try:
                count = archive_dispatched()
                st.success(f"🚚 {count} pedido(s) EM ROTA movidos para o Histórico!")
                st.session_state.show_archive_confirm = False
                trigger_refresh()
                st.rerun()
            except Exception as e:
                st.error(f"Erro ao arquivar: {e}")
    with c2:
        if st.button("Cancelar", icon=":material/close:"):
            st.session_state.show_archive_confirm = False
            st.rerun()

# ══════════════════════════════════════════════════════════════
# NEW PEDIDO FORM
# ══════════════════════════════════════════════════════════════
if st.session_state.show_new_form:
    st.markdown("### Novo Pedido")

    with st.form("new_pedido_form", clear_on_submit=True):
        fc1, fc2, fc3 = st.columns(3)
        with fc1:
            new_pedido = st.text_input("Número do Pedido *", placeholder="Ex: 096075")
        with fc2:
            new_cliente = st.text_input("Cliente *", placeholder="Ex: BIOMA COMERCIO")
        with fc3:
            new_data = st.date_input("Data", value=datetime.now())

        fc4, fc5, fc6 = st.columns(3)
        with fc4:
            new_empresa = st.selectbox("Empresa/Transportadora", EMPRESA_OPTIONS)
        with fc5:
            new_status = st.selectbox("Status", STATUS_OPTIONS, index=0)
        with fc6:
            new_endereco = st.text_input("Endereço", placeholder="Rua, número, cidade/UF")

        fc7, fc8 = st.columns([1, 2])
        with fc7:
            new_nf = st.text_input("Nota Fiscal", placeholder="Número da NF (opcional)")
        with fc8:
            new_obs = st.text_area("Observações", placeholder="Observações adicionais...", height=68)

        submitted = st.form_submit_button("Inserir Pedido", type="primary", width="stretch")

        if submitted:
            if not new_pedido or not new_cliente:
                st.error("Os campos Pedido e Cliente são obrigatórios.")
            else:
                try:
                    insert_pedido({
                        "DATA": datetime.combine(new_data, datetime.min.time()),
                        "CLIENTE": new_cliente.upper().strip(),
                        "PEDIDO": new_pedido.strip(),
                        "EMPRESA": new_empresa,
                        "STATUS": new_status,
                        "OBS": new_obs.strip(),
                        "ENDERECO": new_endereco.strip(),
                        "NF": new_nf.strip(),
                    })
                    st.success(f"Pedido #{new_pedido} inserido com sucesso!")
                    st.session_state.show_new_form = False
                    trigger_refresh()
                    st.rerun()
                except ValueError as e:
                    st.error(f"{e}")
                except Exception as e:
                    st.error(f"Erro ao inserir: {e}")

    st.markdown("---")

# ══════════════════════════════════════════════════════════════
# FILTER DATA
# ══════════════════════════════════════════════════════════════
filtered_df = df.copy()

# Remove pedidos inválidos/vazios para não quebrar contagens ou exibição
filtered_df = filtered_df[
    filtered_df["PEDIDO"].notna() & 
    (filtered_df["PEDIDO"].astype(str).str.strip() != "") & 
    (filtered_df["PEDIDO"].astype(str).str.lower() != "nan") &
    (filtered_df["PEDIDO"].astype(str).str.lower() != "none") &
    (filtered_df["PEDIDO"].astype(str).str.lower() != "<na>")
].copy()

if search_query:
    query_upper = search_query.upper().strip()
    mask = (
        filtered_df["PEDIDO"].astype(str).str.contains(query_upper, case=False, na=False) |
        filtered_df["CLIENTE"].astype(str).str.contains(query_upper, case=False, na=False)
    )
    filtered_df = filtered_df[mask]

# ══════════════════════════════════════════════════════════════
# STATUS FILTER TABS (com separador visual)
# ══════════════════════════════════════════════════════════════
tab_labels = ["Todos"] + [f"{s} ({len(filtered_df[filtered_df['STATUS'] == s])})" for s in STATUS_OPTIONS]
tabs = st.tabs(tab_labels)

# CSS para separador entre tabs + tamanho compacto
st.markdown("""
<style>
    .stTabs [data-baseweb="tab-list"] {
        gap: 0px;
        flex-wrap: nowrap;
    }
    .stTabs [data-baseweb="tab"] {
        border-right: 1px solid rgba(99, 102, 241, 0.2);
        padding: 6px 8px !important;
        font-size: 12px !important;
        white-space: nowrap;
    }
    .stTabs [data-baseweb="tab"]:last-child {
        border-right: none;
    }
    .stTabs [data-baseweb="tab-highlight"] {
        background-color: #6366f1 !important;
    }
</style>
""", unsafe_allow_html=True)

for tab_idx, tab in enumerate(tabs):
    with tab:
        if tab_idx == 0:
            display_df = filtered_df
        else:
            target_status = STATUS_OPTIONS[tab_idx - 1]
            display_df = filtered_df[filtered_df["STATUS"] == target_status]

        if display_df.empty:
            st.markdown("""
            <div style="text-align: center; padding: 40px; color: var(--gb-text-muted, #64748b);">
                <svg xmlns="http://www.w3.org/2000/svg" width="48" height="48" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round" style="margin-bottom: 12px; opacity: 0.5;">
                    <path d="M21 8a2 2 0 0 0-1-1.73l-7-4a2 2 0 0 0-2 0l-7 4A2 2 0 0 0 3 8v8a2 2 0 0 0 1 1.73l7 4a2 2 0 0 0 2 0l7-4A2 2 0 0 0 21 16Z"/>
                    <path d="m3.3 7 8.7 5 8.7-5"/>
                    <path d="M12 22V12"/>
                </svg>
                <p>Nenhum pedido encontrado</p>
            </div>
            """, unsafe_allow_html=True)
        else:
            # ── Preparar DataFrame para edição ──
            edit_df = display_df.copy()
            edit_df["DATA"] = pd.to_datetime(edit_df["DATA"], errors="coerce")
            
            # Limpar valores NaN/None para exibição
            for col in ["CLIENTE", "PEDIDO", "EMPRESA", "STATUS", "OBS", "ENDERECO", "NF"]:
                edit_df[col] = edit_df[col].fillna("").astype(str)
                edit_df[col] = edit_df[col].apply(
                    lambda x: "" if str(x).lower() in ("nan", "none", "<na>") else x
                )

            # Aplicar emojis coloridos no STATUS para visualização
            edit_df["STATUS"] = edit_df["STATUS"].apply(_status_to_display)

            # Adicionar coluna de seleção para exclusão
            edit_df.insert(0, "🗑️", False)

            st.markdown(
                f"<p style='color: var(--gb-text-muted, #64748b); font-size: 13px; margin-bottom: 8px;'>"
                f"Exibindo {len(edit_df)} pedido(s) — edite diretamente na tabela</p>",
                unsafe_allow_html=True
            )

            # ── Tabela editável (inline, como Excel) ──
            edited_df = st.data_editor(
                edit_df,
                use_container_width=True,
                hide_index=True,
                num_rows="fixed",
                height=min(len(edit_df) * 38 + 50, 600),
                key=f"editor_{tab_idx}",
                disabled=not is_adm(),
                column_config={
                    "🗑️": st.column_config.CheckboxColumn(
                        "🗑️",
                        help="Selecione para excluir",
                        width="small",
                        default=False,
                    ),
                    "DATA": st.column_config.DateColumn(
                        "📅 Data",
                        format="DD/MM/YYYY",
                        width="small",
                    ),
                    "CLIENTE": st.column_config.TextColumn(
                        "👤 Cliente",
                        width="large",
                    ),
                    "PEDIDO": st.column_config.TextColumn(
                        "📋 Pedido",
                        width="small",
                        disabled=True,  # PK não editável
                    ),
                    "EMPRESA": st.column_config.SelectboxColumn(
                        "🏢 Empresa",
                        options=EMPRESA_OPTIONS,
                        width="small",
                    ),
                    "STATUS": st.column_config.SelectboxColumn(
                        "🔖 Status",
                        options=STATUS_DISPLAY_OPTIONS,
                        width="medium",
                    ),
                    "OBS": st.column_config.TextColumn(
                        "📝 Obs",
                        width="medium",
                    ),
                    "ENDERECO": st.column_config.TextColumn(
                        "📍 Endereço",
                        width="medium",
                    ),
                    "NF": st.column_config.TextColumn(
                        "📄 NF",
                        width="small",
                    ),
                },
                column_order=["🗑️", "DATA", "CLIENTE", "PEDIDO", "EMPRESA", "NF", "STATUS", "OBS", "ENDERECO"],
            )

            # ── Ações: Salvar + Excluir ──
            act_col1, act_col2, act_col3 = st.columns([1, 1, 3])

            # Detectar mudanças (comparar original vs editado cell-by-cell)
            compare_cols = ["DATA", "CLIENTE", "EMPRESA", "STATUS", "OBS", "ENDERECO", "NF"]
            original_compare = edit_df.drop(columns=["🗑️"]).reset_index(drop=True)
            edited_compare = edited_df.drop(columns=["🗑️"]).reset_index(drop=True)

            # Normalizar para comparação robusta (tudo como string)
            def _normalize_for_compare(val):
                if pd.isna(val) or val is None:
                    return ""
                s = str(val).strip()
                if s.lower() in ("nan", "none", "<na>", "nat"):
                    return ""
                return s

            changed_indices = []
            for idx in range(min(len(original_compare), len(edited_compare))):
                for col in compare_cols:
                    orig_val = _normalize_for_compare(original_compare.iloc[idx].get(col, ""))
                    edit_val = _normalize_for_compare(edited_compare.iloc[idx].get(col, ""))
                    if orig_val != edit_val:
                        changed_indices.append(idx)
                        break

            has_changes = len(changed_indices) > 0

            # Detectar selecionados para exclusão
            selected_for_delete = edited_df[edited_df["🗑️"] == True]
            delete_count = len(selected_for_delete)

            with act_col1:
                if st.button(
                    f"💾 Salvar Alterações",
                    type="primary",
                    use_container_width=True,
                    disabled=not has_changes or not is_adm(),
                    key=f"save_{tab_idx}",
                ):
                    try:
                        # Montar changes apenas com linhas que realmente mudaram
                        changes = []
                        for idx in changed_indices:
                            edit_row = edited_compare.iloc[idx]
                            change = {"PEDIDO": str(edit_row["PEDIDO"]).strip()}
                            for col in ["DATA", "CLIENTE", "EMPRESA", "STATUS", "OBS", "ENDERECO", "NF"]:
                                new_val = edit_row[col]
                                if col == "DATA":
                                    if pd.notna(new_val):
                                        new_val = pd.Timestamp(new_val).to_pydatetime()
                                    else:
                                        new_val = None
                                elif col == "STATUS":
                                    # Strip emoji prefix antes de salvar
                                    new_val = _display_to_status(new_val)
                                elif col == "CLIENTE":
                                    new_val = str(new_val).upper().strip() if pd.notna(new_val) else ""
                                else:
                                    new_val = str(new_val).strip() if pd.notna(new_val) else ""
                                change[col] = new_val
                            changes.append(change)

                        if changes:
                            count = update_rows_batch(changes)
                            st.success(f"✅ {count} pedido(s) atualizado(s)!")
                            trigger_refresh()
                            st.rerun()
                        else:
                            st.info("Nenhuma alteração detectada.")
                    except Exception as e:
                        st.error(f"Erro ao salvar: {e}")

            with act_col2:
                if st.button(
                    f"🗑️ Excluir ({delete_count})",
                    use_container_width=True,
                    disabled=delete_count == 0 or not is_adm(),
                    key=f"delete_{tab_idx}",
                ):
                    st.session_state[f"confirm_batch_del_{tab_idx}"] = True

            # ── Confirmação de exclusão em lote ──
            if st.session_state.get(f"confirm_batch_del_{tab_idx}", False) and delete_count > 0:
                pedidos_to_delete = selected_for_delete["PEDIDO"].astype(str).tolist()
                st.warning(
                    f"⚠️ Confirma exclusão de **{delete_count}** pedido(s)? "
                    f"({', '.join(pedidos_to_delete[:5])}{'...' if len(pedidos_to_delete) > 5 else ''}) "
                    f"**Esta ação não pode ser desfeita.**"
                )
                dc1, dc2, dc3 = st.columns([1, 1, 4])
                with dc1:
                    if st.button("✅ Sim, excluir", type="primary", key=f"cdel_yes_{tab_idx}", use_container_width=True, disabled=not is_adm()):
                        try:
                            count = delete_pedidos_batch(pedidos_to_delete)
                            st.session_state.pop(f"confirm_batch_del_{tab_idx}", None)
                            trigger_refresh()
                            st.success(f"🗑️ {count} pedido(s) excluído(s)!")
                            st.rerun()
                        except Exception as e:
                            st.error(f"Erro: {e}")
                with dc2:
                    if st.button("❌ Cancelar", key=f"cdel_no_{tab_idx}", use_container_width=True):
                        st.session_state.pop(f"confirm_batch_del_{tab_idx}", None)
                        st.rerun()
