"""
8_🚚_Expedicao.py — Módulo de Expedição Diária.

Ciclo completo de expedição:
    Tab 1: 📋 Planejar Expedição — montar plano do dia (atribuir pedidos a veículos)
    Tab 2: 📝 Registrar Viagem — parse de mensagens WhatsApp (migrado da aba 7)
    Tab 3: 📋 Viagens Registradas — CRUD + exportação (migrado da aba 7)
    Tab 4: 🔍 Auditoria — cruzamento WhatsApp × Relatório (migrado da aba 7)
    Tab 5: ✅ Conferência — expedição planejada vs. viagem relatada
    Tab 6: 🔄 Caminho Reverso — confirmação retroativa via NF + relatório
"""

import streamlit as st
from modules.auth import is_adm
import pandas as pd
from datetime import datetime, date, timedelta
from io import BytesIO
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from modules.ui_components import inject_custom_css, render_header
from modules.expedition_engine import (
    get_ready_orders,
    read_expeditions,
    save_expedition,
    delete_expedition,
    get_default_drivers,
    save_default_drivers,
    export_checklist_excel,
    match_nf_to_order,
    cross_check_expedition_vs_trip,
    mark_as_dispatched,
    mark_as_delivered,
    extract_destination,
    get_expedition_dates,
    STATUS_DISPATCHED,
    STATUS_DELIVERED,
)
from modules.whatsapp_parser import (
    parse_single_message, parse_chat_export, format_ocorrencias,
)
from modules.excel_handler import (
    read_viagens, insert_viagem, insert_viagens_batch,
    delete_viagem, update_viagem,
)
from config import FLEET_PLATES, VIAGENS_COLUMNS, EMPRESA_OPTIONS

# Import reverse delivery engine
from modules.reverse_delivery_engine import (
    verify_delivery_in_report,
    confirm_delivery,
    confirm_delivery_batch,
    batch_process_danfes,
    get_pending_orders_without_nf,
)

# Try to import delivery_reader (may not exist in all setups)
try:
    from modules.delivery_reader import (
        read_deliveries_report, get_deliveries_for_plate_date,
    )
    HAS_DELIVERY_REPORT = True
except ImportError:
    HAS_DELIVERY_REPORT = False


# ══════════════════════════════════════════════════════════════
# PAGE CONFIG
# ══════════════════════════════════════════════════════════════
st.set_page_config(page_title="Expedição — Gestão Logística", page_icon="🚚", layout="wide")
inject_custom_css()


# ══════════════════════════════════════════════════════════════
# HELPERS
# ══════════════════════════════════════════════════════════════

def _hex_to_rgb(hex_color):
    h = hex_color.lstrip("#")
    return f"{int(h[0:2], 16)}, {int(h[2:4], 16)}, {int(h[4:6], 16)}"


def _exp_card(label, value, color, icon=""):
    """KPI card for expedition dashboard."""
    return (
        '<div style="'
        f'background: linear-gradient(135deg, rgba({_hex_to_rgb(color)}, 0.10), rgba({_hex_to_rgb(color)}, 0.03));'
        f'border: 1px solid rgba({_hex_to_rgb(color)}, 0.20);'
        f'border-left: 3px solid {color};'
        'border-radius: 10px;'
        'padding: 14px 18px;'
        '">'
        f'<div style="color: {color}; font-size: 11px; font-weight: 700;'
        f' text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 2px;">'
        f'{icon} {label}</div>'
        f'<div style="color: #e2e8f0; font-size: 26px; font-weight: 700;">'
        f'{value}</div>'
        '</div>'
    )


# ══════════════════════════════════════════════════════════════
# HEADER
# ══════════════════════════════════════════════════════════════
render_header("Expedição Diária", "Planejamento de saída, registro de viagens e conferência")

# ══════════════════════════════════════════════════════════════
# TABS
# ══════════════════════════════════════════════════════════════
tab_plan, tab_viagem, tab_viagens_list, tab_audit, tab_check, tab_reverse = st.tabs([
    "📋 Planejar Expedição",
    "📝 Registrar Viagem",
    "📋 Viagens Registradas",
    "🔍 Auditoria",
    "✅ Conferência",
    "🔄 Caminho Reverso",
])


# ══════════════════════════════════════════════════════════════
# TAB 1 — PLANEJAR EXPEDIÇÃO
# ══════════════════════════════════════════════════════════════
with tab_plan:
    # ── Zone 1: Header + Date Selector + KPIs ──
    h1, h2 = st.columns([1, 2])
    with h1:
        exp_date = st.date_input(
            "📅 Data de Saída",
            value=date.today() + timedelta(days=1),
            format="DD/MM/YYYY",
            key="exp_date",
        )

    # Load existing expedition for this date
    df_exp_existing = read_expeditions(exp_date)
    has_existing = not df_exp_existing.empty

    # Initialize session state for assignments
    if "exp_assignments" not in st.session_state:
        st.session_state.exp_assignments = {}
    if "exp_loaded_date" not in st.session_state:
        st.session_state.exp_loaded_date = None

    # Load ready orders (CONCLUIDO / AGUARDANDO NF)
    df_ready = get_ready_orders()

    # Recarrega assignments quando a data muda — busca do arquivo expedicoes.xlsx
    if st.session_state.exp_loaded_date != exp_date:
        if has_existing:
            assignments = {}
            # Pre-computar mapeamento de pedido sem zero -> pedido exato da esteira para padronizar chaves
            ready_pedidos_map = {}
            if not df_ready.empty:
                for p in df_ready["PEDIDO"].astype(str).str.strip():
                    ready_pedidos_map[p.lstrip('0')] = p

            for _, row in df_exp_existing.iterrows():
                pedido = str(row.get("PEDIDO", "")).strip()
                if pedido:
                    # Busca pelo pedido com ou sem zero baseado na esteira
                    pedido_sem_zero = pedido.lstrip('0')
                    if pedido_sem_zero in ready_pedidos_map:
                        pedido = ready_pedidos_map[pedido_sem_zero]
                    else:
                        # Tenta formatar para o padrao 6 digitos caso não exista na esteira mas precise cruzar
                        if pedido.isdigit() and len(pedido) < 6:
                            pedido = pedido.zfill(6)
                            
                    destino_val = str(row.get("DESTINO", "") or "").strip()
                    if destino_val.lower() in ("nan", "none", "0.0", "0"):
                        destino_val = ""
                    assignments[pedido] = {
                        "VEICULO": str(row.get("VEICULO", "")),
                        "ORDEM": int(row.get("ORDEM_ENTREGA", 0)) if pd.notna(row.get("ORDEM_ENTREGA")) else 0,
                        "NF": str(row.get("NF", "") or ""),
                        "PESO": float(row.get("PESO", 0)) if pd.notna(row.get("PESO")) else 0.0,
                        "VOLUMES": int(row.get("VOLUMES", 0)) if pd.notna(row.get("VOLUMES")) else 0,
                        "MOTORISTA": str(row.get("MOTORISTA", "") or ""),
                        "CLIENTE_SALVO": str(row.get("CLIENTE", "") or "").strip(),
                        "DESTINO_MANUAL": destino_val,
                    }
            st.session_state.exp_assignments = assignments
        else:
            # Nova data sem expedição — limpa os assignments
            st.session_state.exp_assignments = {}
        st.session_state.exp_loaded_date = exp_date

    drivers = get_default_drivers()

    # KPIs
    n_ready = len(df_ready)
    n_assigned = len(st.session_state.exp_assignments)
    veiculos_usados = len(set(
        a["VEICULO"] for a in st.session_state.exp_assignments.values()
        if a.get("VEICULO")
    ))

    with h2:
        k1, k2, k3, k4 = st.columns(4)
        with k1:
            st.markdown(
                _exp_card("Prontos", str(n_ready), "#6366f1", "📦"),
                unsafe_allow_html=True,
            )
        with k2:
            st.markdown(
                _exp_card("Atribuídos", str(n_assigned), "#22d3ee", "✅"),
                unsafe_allow_html=True,
            )
        with k3:
            st.markdown(
                _exp_card("Veículos", str(veiculos_usados), "#f59e0b", "🚚"),
                unsafe_allow_html=True,
            )
        with k4:
            st.markdown(
                _exp_card("Pendentes", str(max(0, n_ready - n_assigned)), "#ef4444", "⏳"),
                unsafe_allow_html=True,
            )

    if has_existing:
        st.info(f"📋 Expedição existente para {exp_date.strftime('%d/%m/%Y')} carregada ({len(df_exp_existing)} pedidos).")

    st.markdown("---")

    # ── Zone 2: Planning — Pool + Assignment ──
    col_pool, col_vehicles = st.columns([1, 2])

    with col_pool:
        st.markdown("### 📦 Pedidos Disponíveis")

        if df_ready.empty:
            st.info("Nenhum pedido com status CONCLUIDO ou AGUARDANDO NF.")
        else:
            # Filter out already assigned
            assigned_pedidos = set(st.session_state.exp_assignments.keys())
            df_available = df_ready[~df_ready["PEDIDO"].isin(assigned_pedidos)]

            if df_available.empty:
                st.success("✅ Todos os pedidos prontos foram atribuídos!")
            else:
                st.caption(f"{len(df_available)} pedido(s) disponíveis")

                # Selection checkboxes
                selected_pedidos = []
                for idx, row in df_available.iterrows():
                    pedido = str(row["PEDIDO"]).strip()
                    cliente = str(row.get("CLIENTE", "")).strip()
                    empresa = str(row.get("EMPRESA", "")).strip()
                    obs_val = str(row.get("OBS", "")).strip()
                    destino = extract_destination(
                        address=str(row.get("ENDERECO", "")),
                        cliente=cliente,
                        obs=obs_val,
                    )
                    label = f"**{pedido}** — {cliente}"
                    if destino and destino.lower() != "nan":
                        label += f" → 📍 {destino}"
                    elif empresa:
                        label += f" ({empresa})"
                    if st.checkbox(label, key=f"sel_{pedido}"):
                        selected_pedidos.append(pedido)

                # Assignment controls
                if selected_pedidos:
                    st.markdown("---")
                    vehicle_options = FLEET_PLATES + ["TRANSPORTADORA"]
                    assign_vehicle = st.selectbox(
                        "Atribuir ao veículo:",
                        vehicle_options,
                        key="assign_vehicle",
                    )

                    # Se TRANSPORTADORA, mostrar campo para nome
                    final_vehicle = assign_vehicle
                    if assign_vehicle == "TRANSPORTADORA":
                        transport_name = st.text_input(
                            "Nome da transportadora:",
                            placeholder="Ex: BRASPRESS, TNT, JADLOG...",
                            key="transport_name",
                        )
                        if transport_name:
                            final_vehicle = transport_name.upper().strip()
                        else:
                            final_vehicle = ""

                    if st.button("▶ Atribuir Selecionados", type="primary", use_container_width=True, disabled=(not final_vehicle) or not is_adm()):
                        for pedido in selected_pedidos:
                            row_data = df_ready[df_ready["PEDIDO"] == pedido].iloc[0]
                            st.session_state.exp_assignments[pedido] = {
                                "VEICULO": final_vehicle,
                                "ORDEM": len([
                                    a for a in st.session_state.exp_assignments.values()
                                    if a.get("VEICULO") == final_vehicle
                                ]) + 1,
                                "NF": "",
                                "PESO": 0,
                                "VOLUMES": 0,
                                "MOTORISTA": drivers.get(final_vehicle, ""),
                            }
                        st.success(f"✅ {len(selected_pedidos)} pedido(s) atribuídos ao {final_vehicle}")
                        st.rerun()

        # ── NF Upload (optional) ──
        st.markdown("---")
        st.markdown("### 📄 Associar NF")
        nf_mode = st.radio(
            "Modo:", ["Upload PDF", "Manual"],
            horizontal=True,
            key="nf_mode",
            label_visibility="collapsed",
        )

        if nf_mode == "Upload PDF":
            nf_files = st.file_uploader(
                "Upload da NF (DANFE)",
                type=["pdf"],
                key="nf_upload",
                accept_multiple_files=True,
                disabled=not is_adm(),
            )
            if nf_files:
                for nf_file in nf_files:
                    try:
                        from modules.danfe_parser import extrair_danfe
                        nf_file.seek(0)
                        nf_data = extrair_danfe(nf_file)
                        nf_file.seek(0)

                        nf_num = nf_data.get("Nota_Fiscal") or nf_data.get("NF", "")
                        if nf_data and nf_num:
                            st.success(f"NF **{nf_num}** extraída de `{nf_file.name}` — Cliente: {nf_data.get('Cliente', '?')}")

                            # Try auto-match only against orders in current trip that DON'T have an NF yet
                            assigned_no_nf_norm = [
                                str(p).strip().zfill(6) for p, a in st.session_state.exp_assignments.items()
                                if not a.get("NF")
                            ]
                            
                            # Build search pool: unassigned (df_ready) + already assigned in this trip
                            assigned_records = []
                            for p, a in st.session_state.exp_assignments.items():
                                assigned_records.append({
                                    "PEDIDO": p,
                                    "CLIENTE": a.get("CLIENTE_SALVO", ""),
                                    "STATUS": "EM ROTA"
                                })
                            
                            df_assigned = pd.DataFrame(assigned_records)
                            if not df_assigned.empty:
                                df_search = pd.concat([df_ready, df_assigned]).drop_duplicates(subset=["PEDIDO"])
                            else:
                                df_search = df_ready.copy()
                            
                            matched = match_nf_to_order(nf_data, df_search)
                            
                            # Check if it already has an NF (accounting for np.nan and whitespace)
                            existing_nf = ""
                            if matched and matched in st.session_state.exp_assignments:
                                val = st.session_state.exp_assignments[matched].get("NF", "")
                                if pd.notna(val) and str(val).strip() and str(val).strip().lower() != "nan":
                                    existing_nf = str(val).strip()
                            
                            if matched and matched in st.session_state.exp_assignments and not existing_nf:
                                st.session_state.exp_assignments[matched]["NF"] = nf_num
                                peso_str = nf_data.get("Peso") or nf_data.get("Peso", "")
                                if peso_str:
                                    try:
                                        st.session_state.exp_assignments[matched]["PESO"] = float(
                                            str(peso_str).replace(",", ".")
                                        )
                                    except (ValueError, TypeError):
                                        pass
                                vol_str = nf_data.get("Volumes", "")
                                if vol_str:
                                    try:
                                        st.session_state.exp_assignments[matched]["VOLUMES"] = int(vol_str)
                                    except (ValueError, TypeError):
                                        pass
                                st.info(f"✅ NF {nf_num} associada ao pedido **{matched}**")
                            elif matched and matched in st.session_state.exp_assignments and existing_nf:
                                st.warning(f"Pedido {matched} já possui uma NF associada ({existing_nf}). Remova-a antes de associar uma nova.")
                            elif matched:
                                st.warning(f"Pedido {matched} encontrado mas ainda não atribuído a um veículo.")
                            else:
                                st.warning(f"`{nf_file.name}`: Não foi possível associar automaticamente. Use o modo manual.")
                        else:
                            st.warning(f"Não foi possível extrair dados de `{nf_file.name}`.")
                    except Exception as e:
                        st.error(f"Erro ao processar `{nf_file.name}`: {e}")
        else:
            nf_manual_pedido = st.selectbox(
                "Pedido:",
                list(st.session_state.exp_assignments.keys()) or ["—"],
                key="nf_manual_pedido",
            )
            nf_manual_num = st.text_input("Número da NF:", key="nf_manual_num")
            nf_manual_peso = st.number_input("Peso (kg):", min_value=0.0, step=0.1, key="nf_manual_peso")
            nf_manual_vol = st.number_input("Volumes:", min_value=0, step=1, key="nf_manual_vol")

            if st.button("💾 Associar NF", use_container_width=True, disabled=not is_adm()):
                if nf_manual_pedido in st.session_state.exp_assignments and nf_manual_num:
                    st.session_state.exp_assignments[nf_manual_pedido]["NF"] = nf_manual_num.strip()
                    if nf_manual_peso > 0:
                        st.session_state.exp_assignments[nf_manual_pedido]["PESO"] = nf_manual_peso
                    if nf_manual_vol > 0:
                        st.session_state.exp_assignments[nf_manual_pedido]["VOLUMES"] = nf_manual_vol
                    st.success(f"✅ NF {nf_manual_num} → Pedido {nf_manual_pedido}")
                    st.rerun()

    # ── Vehicle Columns ──
    with col_vehicles:
        st.markdown("### 🚚 Veículos")

        if not st.session_state.exp_assignments:
            st.info("Nenhum pedido atribuído ainda. Selecione pedidos no painel esquerdo.")
        else:
            # Coletar todos os veículos/transportadoras atribuídos
            all_vehicles = list(FLEET_PLATES)
            for a in st.session_state.exp_assignments.values():
                v = a.get("VEICULO", "")
                if v and v not in all_vehicles:
                    all_vehicles.append(v)

            v_cols = st.columns(len(all_vehicles))

            fleet_colors = ["#6366f1", "#22d3ee", "#f59e0b"]
            transport_color = "#a855f7"  # Purple for external carriers

            for v_idx, placa in enumerate(all_vehicles):
                with v_cols[v_idx]:
                    # Get assignments for this vehicle
                    v_pedidos = {
                        p: a for p, a in st.session_state.exp_assignments.items()
                        if a.get("VEICULO") == placa
                    }

                    is_fleet = placa in FLEET_PLATES
                    motorista = drivers.get(placa, "") if is_fleet else ""
                    n_pedidos = len(v_pedidos)

                    # Vehicle header
                    if is_fleet:
                        color = fleet_colors[FLEET_PLATES.index(placa) % len(fleet_colors)]
                        icon = "🚚"
                    else:
                        color = transport_color
                        icon = "📦"

                    subtitle = f"{motorista} — {n_pedidos} pedido(s)" if motorista else f"{n_pedidos} pedido(s)"
                    st.markdown(f"""
                    <div style="
                        background: linear-gradient(135deg, rgba({_hex_to_rgb(color)}, 0.15), rgba({_hex_to_rgb(color)}, 0.05));
                        border: 1px solid rgba({_hex_to_rgb(color)}, 0.25);
                        border-radius: 10px;
                        padding: 10px 14px;
                        margin-bottom: 8px;
                    ">
                        <div style="font-weight: 700; color: {color}; font-size: 14px;">{icon} {placa}</div>
                        <div style="color: #94a3b8; font-size: 12px;">{subtitle}</div>
                    </div>
                    """, unsafe_allow_html=True)

                    if v_pedidos:
                        for pedido, assign in sorted(v_pedidos.items(), key=lambda x: x[1].get("ORDEM", 99)):
                            # Get order details
                            row_match = df_ready[df_ready["PEDIDO"] == pedido]
                            cliente = assign.get("CLIENTE_SALVO", "")
                            if not cliente:
                                cliente = str(row_match.iloc[0]["CLIENTE"]) if not row_match.empty else "?"
                            nf_val = assign.get("NF", "")

                            st.markdown(f"""
                            <div style="
                                background: rgba(255,255,255,0.03);
                                border: 1px solid rgba(255,255,255,0.06);
                                border-radius: 8px;
                                padding: 8px 10px;
                                margin-bottom: 4px;
                                font-size: 12px;
                            ">
                                <div style="color: #e2e8f0; font-weight: 600;">#{assign.get('ORDEM', '—')} — {pedido}</div>
                                <div style="color: #94a3b8;">{cliente[:25]}</div>
                                <div style="color: #64748b; font-size: 11px;">
                                    NF: {nf_val or '—'} | {assign.get('PESO', 0):.0f}kg | {assign.get('VOLUMES', 0)} vol
                                </div>
                            </div>
                            """, unsafe_allow_html=True)

                            # Order and remove controls
                            oc1, oc2 = st.columns([2, 1])
                            with oc1:
                                new_order = st.number_input(
                                    "Ord",
                                    value=max(assign.get("ORDEM", 1), 1),
                                    min_value=1,
                                    step=1,
                                    key=f"ord_{placa}_{pedido}",
                                    label_visibility="collapsed",
                                )
                                if new_order != assign.get("ORDEM"):
                                    st.session_state.exp_assignments[pedido]["ORDEM"] = new_order
                            with oc2:
                                if st.button("✖", key=f"rem_{placa}_{pedido}", help="Remover"):
                                    del st.session_state.exp_assignments[pedido]
                                    st.rerun()
                    else:
                        st.caption("Nenhum pedido atribuído")

    # ── Zone 3: Actions ──
    st.markdown("---")

    if st.session_state.exp_assignments:
        act1, act2, act3, act4 = st.columns(4)

        with act1:
            if st.button("💾 Salvar Plano", type="primary", use_container_width=True, disabled=not is_adm()):
                items = []
                for pedido, assign in st.session_state.exp_assignments.items():
                    row_match = df_ready[df_ready["PEDIDO"] == pedido]
                    cliente = assign.get("CLIENTE_SALVO", "")
                    if not cliente:
                        cliente = str(row_match.iloc[0]["CLIENTE"]) if not row_match.empty else ""
                    endereco = str(row_match.iloc[0].get("ENDERECO", "")) if not row_match.empty else ""
                    obs_text = str(row_match.iloc[0].get("OBS", "")) if not row_match.empty else ""

                    items.append({
                        "DATA_EXPEDICAO": exp_date,
                        "PEDIDO": pedido,
                        "CLIENTE": cliente,
                        "DESTINO": assign.get("DESTINO_MANUAL") if assign.get("DESTINO_MANUAL") is not None else extract_destination(
                            address=endereco,
                            cliente=cliente,
                            obs=obs_text,
                        ),
                        "VEICULO": assign.get("VEICULO", ""),
                        "MOTORISTA": assign.get("MOTORISTA", drivers.get(assign.get("VEICULO", ""), "")),
                        "ORDEM_ENTREGA": assign.get("ORDEM", 0),
                        "NF": assign.get("NF", ""),
                        "PESO": assign.get("PESO", 0),
                        "VOLUMES": assign.get("VOLUMES", 0),
                        "CARREGADO": False,
                        "OBS": "",
                        "CRIADO_EM": datetime.now(),
                    })

                count = save_expedition(items)
                st.success(f"✅ Plano salvo com {count} pedido(s)!")

                # Mark orders as EM ROTA
                n_updated = mark_as_dispatched(exp_date)
                if n_updated > 0:
                    st.info(f"🚚 {n_updated} pedido(s) marcados como **EM ROTA** na Esteira.")

                st.session_state.exp_loaded_date = exp_date

        with act2:
            if st.button("📥 Exportar Checklist", use_container_width=True, disabled=not is_adm()):
                try:
                    # Save first if not saved
                    df_check = read_expeditions(exp_date)
                    if df_check.empty:
                        st.warning("⚠️ Salve o plano antes de exportar.")
                    else:
                        excel_bytes = export_checklist_excel(exp_date)
                        st.download_button(
                            label="⬇️ Download Checklist Excel",
                            data=excel_bytes,
                            file_name=f"checklist_{exp_date.strftime('%d%m%Y')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True,
                        )
                except Exception as e:
                    st.error(f"Erro ao exportar: {e}")

        with act3:
            if st.button("🗑️ Limpar Plano", use_container_width=True, disabled=not is_adm()):
                st.session_state["confirm_clear_exp"] = True

            if st.session_state.get("confirm_clear_exp"):
                st.warning("⚠️ Tem certeza que deseja limpar todas as atribuições?")
                cc1, cc2 = st.columns(2)
                with cc1:
                    if st.button("✅ Sim, limpar", key="confirm_clear_yes", disabled=not is_adm()):
                        st.session_state.exp_assignments = {}
                        st.session_state.exp_loaded_date = None
                        st.session_state.pop("confirm_clear_exp", None)
                        delete_expedition(exp_date)
                        st.rerun()
                with cc2:
                    if st.button("❌ Cancelar", key="confirm_clear_no"):
                        st.session_state.pop("confirm_clear_exp", None)
                        st.rerun()

        with act4:
            # Driver config
            with st.popover("⚙️ Motoristas Padrão"):
                st.markdown("#### Motoristas por Veículo")
                new_drivers = {}
                for placa in FLEET_PLATES:
                    new_drivers[placa] = st.text_input(
                        placa,
                        value=drivers.get(placa, ""),
                        key=f"drv_{placa}",
                    )
                if st.button("💾 Salvar", key="save_drivers", disabled=not is_adm()):
                    save_default_drivers(new_drivers)
                    st.success("Motoristas atualizados!")
                    st.rerun()

    # ── Preview table ──
    if st.session_state.exp_assignments:
        st.markdown("### 👁️ Preview do Checklist")

        st.caption("Você pode alterar Destino, NF, Peso e Volumes diretamente na tabela abaixo. Clique em **Aplicar Alterações** antes de Salvar.")

        preview_data = []
        for pedido, assign in st.session_state.exp_assignments.items():
            row_match = df_ready[df_ready["PEDIDO"] == pedido]
            cliente = assign.get("CLIENTE_SALVO", "")
            if not cliente:
                cliente = str(row_match.iloc[0]["CLIENTE"]) if not row_match.empty else ""
            obs_val = str(row_match.iloc[0].get("OBS", "")) if not row_match.empty else ""
            endereco_val = str(row_match.iloc[0].get("ENDERECO", "")) if not row_match.empty else ""
            
            destino = assign.get("DESTINO_MANUAL")
            if not destino:
                destino = extract_destination(address=endereco_val, cliente=cliente, obs=obs_val)
                if str(destino).lower() in ("nan", "none", "0.0", "0"):
                    destino = ""

            nf_val = str(assign.get("NF", "")).strip()
            if nf_val.lower() in ("nan", "none", "0.0", "0"):
                nf_val = ""
            elif nf_val.endswith(".0"):
                nf_val = nf_val[:-2]

            peso_val = pd.to_numeric(assign.get("PESO", 0), errors="coerce")
            peso_val = peso_val if not pd.isna(peso_val) else 0

            vol_val = pd.to_numeric(assign.get("VOLUMES", 0), errors="coerce")
            vol_val = int(vol_val) if not pd.isna(vol_val) else 0

            peso_str = f"{peso_val:g}".replace(".", ",") if peso_val else ""

            preview_data.append({
                "Veículo": assign.get("VEICULO", ""),
                "Ord": assign.get("ORDEM", 0),
                "Pedido": pedido,
                "Cliente": cliente,
                "Destino": destino,
                "NF": nf_val,
                "Peso (kg)": peso_str,
                "Volumes": str(vol_val) if vol_val else "",
            })

        df_preview = pd.DataFrame(preview_data)
        df_preview = df_preview.sort_values(["Veículo", "Ord"]).reset_index(drop=True)
        
        edited_df = st.data_editor(
            df_preview,
            use_container_width=True,
            hide_index=True,
            key="preview_editor",
            disabled=["Veículo", "Ord", "Pedido", "Cliente"]
        )
        
        if st.button("✏️ Aplicar Alterações da Tabela", type="secondary", disabled=not is_adm()):
            for idx, row in edited_df.iterrows():
                pedido_row = str(row["Pedido"])
                if pedido_row in st.session_state.exp_assignments:
                    st.session_state.exp_assignments[pedido_row]["DESTINO_MANUAL"] = str(row["Destino"]).strip()
                    st.session_state.exp_assignments[pedido_row]["NF"] = str(row["NF"]).strip()
                    
                    peso = str(row["Peso (kg)"]).strip().replace(",", ".")
                    vol = str(row["Volumes"]).strip()
                    
                    try:
                        st.session_state.exp_assignments[pedido_row]["PESO"] = float(peso) if peso else 0
                    except:
                        pass

                    try:
                        st.session_state.exp_assignments[pedido_row]["VOLUMES"] = int(float(vol)) if vol else 0
                    except:
                        pass
            
            st.success("Alterações aplicadas! Você já pode clicar em Salvar Plano.")
            st.rerun()


# ══════════════════════════════════════════════════════════════
# TAB 2 — REGISTRAR VIAGEM (migrado da aba 7)
# ══════════════════════════════════════════════════════════════
with tab_viagem:
    st.markdown("### 📝 Registrar Viagem do WhatsApp")

    mode = st.radio(
        "Modo de entrada:",
        ["Colar mensagem individual", "Importar chat (.txt)"],
        horizontal=True,
        key="exp_input_mode",
    )

    if mode == "Colar mensagem individual":
        msg_text = st.text_area(
            "Cole a mensagem do WhatsApp aqui:",
            height=180,
            placeholder=(
                "QJJ 9302\nJefferson\nPedido 57\nEntregue 51\n"
                "4 não deu tempo Sabará\n2 devolveu brf\n"
                "Km inicial 316096\nKm final 316152"
            ),
            key="exp_msg_input",
        )

        if msg_text and msg_text.strip():
            parsed = parse_single_message(msg_text)

            # Show warnings
            for w in parsed.get("warnings", []):
                st.warning(f"⚠️ {w}")

            # Check if there's an expedition for this plate+date
            if parsed.get("placa"):
                check_data = parsed["data"].date() if hasattr(parsed["data"], 'date') else date.today()
                exp_match = read_expeditions(check_data)
                exp_for_plate = exp_match[exp_match["VEICULO"] == parsed["placa"]] if not exp_match.empty else pd.DataFrame()
                if not exp_for_plate.empty:
                    st.markdown(f"""
                    <div style="
                        background: rgba(34, 211, 238, 0.08);
                        border: 1px solid rgba(34, 211, 238, 0.25);
                        border-radius: 10px;
                        padding: 12px 16px;
                        margin-bottom: 12px;
                    ">
                        <div style="color: #22d3ee; font-weight: 700; font-size: 13px;">
                            📋 Expedição encontrada para {parsed['placa']} em {check_data.strftime('%d/%m/%Y')}
                        </div>
                        <div style="color: #94a3b8; font-size: 12px;">
                            {len(exp_for_plate)} pedidos planejados — A conferência será habilitada na aba ✅
                        </div>
                    </div>
                    """, unsafe_allow_html=True)

            st.markdown("#### Dados Extraídos")

            with st.form("viagem_form_exp", clear_on_submit=False):
                fc1, fc2, fc3, fc4 = st.columns(4)
                with fc1:
                    v_data = st.date_input(
                        "Data",
                        value=parsed["data"].date() if hasattr(parsed["data"], 'date') else date.today(),
                        format="DD/MM/YYYY",
                    )
                with fc2:
                    v_motorista = st.text_input("Motorista", value=parsed["motorista"])
                with fc3:
                    v_ajudante = st.text_input("Ajudante", value=parsed.get("ajudante", ""))
                with fc4:
                    plate_options = FLEET_PLATES + ["Outra"]
                    default_idx = 0
                    if parsed["placa"] in FLEET_PLATES:
                        default_idx = FLEET_PLATES.index(parsed["placa"])
                    v_placa = st.selectbox("Placa", plate_options, index=default_idx)

                fc5, fc6, fc7, fc8 = st.columns(4)
                with fc5:
                    v_pedidos = st.number_input("Pedidos", value=parsed["pedidos"], min_value=0)
                with fc6:
                    v_entregas = st.number_input("Entregas", value=parsed["entregas"], min_value=0)
                with fc7:
                    v_coletas = st.number_input("Coletas", value=parsed["coletas"], min_value=0)
                with fc8:
                    v_km_ini = st.number_input("KM Inicial", value=parsed["km_inicial"], min_value=0)

                fc9, fc10 = st.columns(2)
                with fc9:
                    v_km_fim = st.number_input("KM Final", value=parsed["km_final"], min_value=0)
                with fc10:
                    km_rodado = v_km_fim - v_km_ini if v_km_fim > v_km_ini else 0
                    st.metric("KM Rodado", f"{km_rodado} km")

                v_ocorrencias = st.text_input(
                    "Ocorrências",
                    value=format_ocorrencias(parsed["ocorrencias"]),
                )
                v_obs = st.text_area("Observações", value=parsed["obs"], height=80)

                submitted = st.form_submit_button("💾 Salvar Viagem", type="primary", use_container_width=True)

                if submitted:
                    viagem = {
                        "DATA": datetime.combine(v_data, datetime.min.time()),
                        "MOTORISTA": v_motorista,
                        "AJUDANTE": v_ajudante,
                        "PLACA": v_placa,
                        "PEDIDOS": v_pedidos,
                        "ENTREGAS": v_entregas,
                        "COLETAS": v_coletas,
                        "KM_INICIAL": v_km_ini,
                        "KM_FINAL": v_km_fim,
                        "KM_RODADO": km_rodado,
                        "OCORRENCIAS": v_ocorrencias,
                        "OBS": v_obs,
                        "TEXTO_ORIGINAL": msg_text.strip(),
                    }
                    try:
                        insert_viagem(viagem)
                        st.success("✅ Viagem salva com sucesso!")
                        st.balloons()
                    except Exception as e:
                        st.error(f"Erro ao salvar: {e}")

    else:  # Import .txt
        uploaded = st.file_uploader(
            "Upload do chat exportado (.txt)",
            type=["txt"],
            key="exp_chat_upload",
            disabled=not is_adm(),
        )

        if uploaded:
            content = uploaded.read().decode("utf-8", errors="replace")
            parsed_msgs = parse_chat_export(content)

            if not parsed_msgs:
                st.warning("Nenhuma mensagem de viagem encontrada no arquivo.")
            else:
                st.success(f"🔍 {len(parsed_msgs)} viagens encontradas!")

                preview_data = []
                for p in parsed_msgs:
                    preview_data.append({
                        "Data": p["data"].strftime("%d/%m/%Y") if hasattr(p["data"], 'strftime') else "",
                        "Motorista": p["motorista"],
                        "Placa": p["placa"],
                        "Pedidos": p["pedidos"],
                        "Entregas": p["entregas"],
                        "Coletas": p["coletas"],
                        "KM": p["km_rodado"],
                    })

                st.dataframe(pd.DataFrame(preview_data), use_container_width=True, hide_index=True)

                if st.button("💾 Salvar Todas as Viagens", type="primary", use_container_width=True, key="exp_save_batch", disabled=not is_adm()):
                    batch = []
                    for p in parsed_msgs:
                        batch.append({
                            "DATA": p["data"],
                            "MOTORISTA": p["motorista"],
                            "AJUDANTE": p.get("ajudante", ""),
                            "PLACA": p["placa"],
                            "PEDIDOS": p["pedidos"],
                            "ENTREGAS": p["entregas"],
                            "COLETAS": p["coletas"],
                            "KM_INICIAL": p["km_inicial"],
                            "KM_FINAL": p["km_final"],
                            "KM_RODADO": p["km_rodado"],
                            "OCORRENCIAS": p["ocorrencias"],
                            "OBS": p["obs"],
                            "TEXTO_ORIGINAL": p["texto_original"],
                        })
                    try:
                        count = insert_viagens_batch(batch)
                        st.success(f"✅ {count} viagens salvas com sucesso!")
                        st.balloons()
                    except Exception as e:
                        st.error(f"Erro ao salvar: {e}")


# ══════════════════════════════════════════════════════════════
# TAB 3 — VIAGENS REGISTRADAS (migrado da aba 7)
# ══════════════════════════════════════════════════════════════
with tab_viagens_list:
    st.markdown("### 📋 Viagens Registradas")

    df_viag = read_viagens()
    if df_viag.empty:
        st.info("Nenhuma viagem registrada ainda.")
    else:
        df_sorted = df_viag.sort_values("DATA", ascending=False)
        display_cols = [c for c in df_sorted.columns if c != "TEXTO_ORIGINAL"]
        st.dataframe(df_sorted[display_cols], use_container_width=True, hide_index=False)

        # Export
        export_df = df_sorted[display_cols].copy()
        if "DATA" in export_df.columns:
            export_df["DATA"] = pd.to_datetime(export_df["DATA"], errors="coerce").dt.strftime("%d/%m/%Y")
        buffer = BytesIO()
        export_df.to_excel(buffer, index=False, sheet_name="Viagens", engine="openpyxl")
        buffer.seek(0)
        st.download_button(
            label="📥 Exportar Viagens para Excel",
            data=buffer,
            file_name=f"viagens_export_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

        # Management
        st.markdown("#### ⚙️ Gerenciar Viagem")
        st.caption("Use o **índice** da tabela acima para selecionar a viagem.")

        mgmt_col1, mgmt_col2 = st.columns(2)

        with mgmt_col1:
            selected_idx = st.number_input(
                "Índice da viagem",
                min_value=0,
                max_value=max(0, len(df_viag) - 1),
                value=0,
                step=1,
                key="exp_mgmt_viagem_idx",
            )

        if selected_idx in df_viag.index:
            sel_row = df_viag.loc[selected_idx]
            sel_date = sel_row.get('DATA', '')
            if hasattr(sel_date, 'strftime'):
                sel_date = sel_date.strftime('%d/%m/%Y')
            st.info(
                f"**Selecionada:** {sel_date} — "
                f"{sel_row.get('MOTORISTA', '')} — "
                f"{sel_row.get('PLACA', '')} — "
                f"{sel_row.get('PEDIDOS', 0)} pedidos"
            )

        with mgmt_col2:
            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("🗑️ Excluir Viagem", type="secondary", key="exp_btn_delete_viagem", use_container_width=True, disabled=not is_adm()):
                st.session_state["exp_confirm_delete_viagem"] = selected_idx

        if st.session_state.get("exp_confirm_delete_viagem") is not None:
            del_idx = st.session_state["exp_confirm_delete_viagem"]
            del_row = df_viag.loc[del_idx] if del_idx in df_viag.index else None
            if del_row is not None:
                del_date = del_row.get('DATA', '')
                if hasattr(del_date, 'strftime'):
                    del_date = del_date.strftime('%d/%m/%Y')
                st.warning(
                    f"⚠️ Confirma exclusão da viagem **#{del_idx}** "
                    f"({del_date} — {del_row.get('MOTORISTA', '')} — {del_row.get('PLACA', '')})?"
                )
                dc1, dc2 = st.columns(2)
                with dc1:
                    if st.button("✅ Sim, excluir", type="primary", key="exp_confirm_del_yes", use_container_width=True, disabled=not is_adm()):
                        try:
                            delete_viagem(del_idx)
                            st.session_state.pop("exp_confirm_delete_viagem", None)
                            st.success("Viagem excluída com sucesso!")
                            st.rerun()
                        except Exception as e:
                            st.error(f"Erro: {e}")
                with dc2:
                    if st.button("❌ Cancelar", key="exp_confirm_del_no", use_container_width=True):
                        st.session_state.pop("exp_confirm_delete_viagem", None)
                        st.rerun()

        if st.button("✏️ Editar Viagem Selecionada", key="exp_btn_edit_viagem", use_container_width=True, disabled=not is_adm()):
            st.session_state["exp_editing_viagem_idx"] = selected_idx

        if st.session_state.get("exp_editing_viagem_idx") is not None:
            edit_idx = st.session_state["exp_editing_viagem_idx"]
            if edit_idx in df_viag.index:
                edit_row = df_viag.loc[edit_idx]
                st.markdown(f"##### ✏️ Editando viagem **#{edit_idx}**")

                with st.form("exp_edit_viagem_form", clear_on_submit=False):
                    ec1, ec2, ec3, ec4 = st.columns(4)
                    with ec1:
                        e_data = st.date_input(
                            "Data",
                            value=edit_row["DATA"].date() if hasattr(edit_row.get("DATA"), 'date') else date.today(),
                            format="DD/MM/YYYY",
                            key="exp_edit_data",
                        )
                    with ec2:
                        e_motorista = st.text_input("Motorista", value=str(edit_row.get("MOTORISTA", "")), key="exp_edit_mot")
                    with ec3:
                        e_ajudante = st.text_input("Ajudante", value=str(edit_row.get("AJUDANTE", "") or ""), key="exp_edit_aju")
                    with ec4:
                        plate_options = FLEET_PLATES + ["Outra"]
                        cur_placa = str(edit_row.get("PLACA", ""))
                        p_idx = plate_options.index(cur_placa) if cur_placa in plate_options else 0
                        e_placa = st.selectbox("Placa", plate_options, index=p_idx, key="exp_edit_placa")

                    ec5, ec6, ec7, ec8 = st.columns(4)
                    with ec5:
                        e_pedidos = st.number_input("Pedidos", value=int(edit_row.get("PEDIDOS", 0) or 0), min_value=0, key="exp_edit_ped")
                    with ec6:
                        e_entregas = st.number_input("Entregas", value=int(edit_row.get("ENTREGAS", 0) or 0), min_value=0, key="exp_edit_ent")
                    with ec7:
                        e_coletas = st.number_input("Coletas", value=int(edit_row.get("COLETAS", 0) or 0), min_value=0, key="exp_edit_col")
                    with ec8:
                        e_km_ini = st.number_input("KM Inicial", value=int(edit_row.get("KM_INICIAL", 0) or 0), min_value=0, key="exp_edit_kmi")

                    ec9, ec10 = st.columns(2)
                    with ec9:
                        e_km_fim = st.number_input("KM Final", value=int(edit_row.get("KM_FINAL", 0) or 0), min_value=0, key="exp_edit_kmf")
                    with ec10:
                        e_km_rod = e_km_fim - e_km_ini if e_km_fim > e_km_ini else 0
                        st.metric("KM Rodado", f"{e_km_rod} km")

                    e_ocorrencias = st.text_input("Ocorrências", value=str(edit_row.get("OCORRENCIAS", "") or ""), key="exp_edit_ocr")
                    e_obs = st.text_area("Observações", value=str(edit_row.get("OBS", "") or ""), height=80, key="exp_edit_obs")

                    ef1, ef2 = st.columns(2)
                    with ef1:
                        save_edit = st.form_submit_button("💾 Salvar Alterações", type="primary", use_container_width=True, disabled=not is_adm())
                    with ef2:
                        cancel_edit = st.form_submit_button("❌ Cancelar Edição", use_container_width=True)

                    if save_edit:
                        updated = {
                            "DATA": datetime.combine(e_data, datetime.min.time()),
                            "MOTORISTA": e_motorista,
                            "AJUDANTE": e_ajudante,
                            "PLACA": e_placa,
                            "PEDIDOS": e_pedidos,
                            "ENTREGAS": e_entregas,
                            "COLETAS": e_coletas,
                            "KM_INICIAL": e_km_ini,
                            "KM_FINAL": e_km_fim,
                            "KM_RODADO": e_km_rod,
                            "OCORRENCIAS": e_ocorrencias,
                            "OBS": e_obs,
                        }
                        try:
                            update_viagem(edit_idx, updated)
                            st.session_state.pop("exp_editing_viagem_idx", None)
                            st.success("✅ Viagem atualizada com sucesso!")
                            st.rerun()
                        except Exception as e:
                            st.error(f"Erro: {e}")

                    if cancel_edit:
                        st.session_state.pop("exp_editing_viagem_idx", None)
                        st.rerun()


# ══════════════════════════════════════════════════════════════
# TAB 4 — AUDITORIA (migrado da aba 7)
# ══════════════════════════════════════════════════════════════
with tab_audit:
    st.markdown("### 🔍 Auditoria — Cruzamento WhatsApp × Relatório")

    df_viag_audit = read_viagens()

    if not HAS_DELIVERY_REPORT:
        st.warning("Módulo de relatório de entregas não disponível.")
    elif df_viag_audit.empty:
        st.info("Registre viagens na aba anterior para habilitar a auditoria.")
    else:
        try:
            df_del = read_deliveries_report()
        except Exception:
            df_del = pd.DataFrame()

        if df_del.empty:
            st.warning("Relatório de entregas não disponível.")
        else:
            ac1, ac2 = st.columns(2)
            with ac1:
                audit_placa = st.selectbox("Placa", FLEET_PLATES, key="exp_audit_placa")
            with ac2:
                audit_periodo = st.date_input(
                    "Período", value=(date(2026, 1, 1), date.today()),
                    key="exp_audit_periodo",
                )

            v_mask = df_viag_audit["PLACA"] == audit_placa
            if isinstance(audit_periodo, tuple) and len(audit_periodo) == 2:
                df_viag_audit["DATA"] = pd.to_datetime(df_viag_audit["DATA"], errors="coerce")
                v_mask = v_mask & (df_viag_audit["DATA"].dt.date >= audit_periodo[0])
                v_mask = v_mask & (df_viag_audit["DATA"].dt.date <= audit_periodo[1])

            viagens_filtered = df_viag_audit[v_mask].copy()

            if viagens_filtered.empty:
                st.info(f"Nenhuma viagem registrada para {audit_placa} no período.")
            else:
                st.markdown(f"**{len(viagens_filtered)} viagens** de `{audit_placa}`")

                audit_rows = []
                for _, viagem in viagens_filtered.iterrows():
                    v_date = viagem["DATA"]
                    if pd.isna(v_date):
                        continue

                    nfs = get_deliveries_for_plate_date(df_del, audit_placa, v_date)
                    nfs_count = len(nfs)
                    pedidos_wpp = int(viagem.get("PEDIDOS", 0) or 0)
                    entregas_wpp = int(viagem.get("ENTREGAS", 0) or 0)

                    divergencia = abs(nfs_count - pedidos_wpp) if pedidos_wpp > 0 else 0

                    if pedidos_wpp == 0:
                        status = "ℹ️ Sem pedidos informados"
                    elif divergencia <= 1:
                        status = "✅ OK"
                    elif nfs_count > pedidos_wpp:
                        status = (
                            f"⚠️ Relatório tem {nfs_count - pedidos_wpp} NF(s) a mais "
                            f"que o WhatsApp ({nfs_count} NFs vs {pedidos_wpp} pedidos)"
                        )
                    else:
                        status = (
                            f"⚠️ WhatsApp tem {pedidos_wpp - nfs_count} pedido(s) a mais "
                            f"que o Relatório ({pedidos_wpp} pedidos vs {nfs_count} NFs)"
                        )

                    audit_rows.append({
                        "Data": v_date.strftime("%d/%m/%Y") if hasattr(v_date, 'strftime') else "",
                        "Motorista": viagem.get("MOTORISTA", ""),
                        "Pedidos (WhatsApp)": pedidos_wpp,
                        "Entregas (WhatsApp)": entregas_wpp,
                        "NFs (Relatório)": nfs_count,
                        "Peso Total (kg)": f"{nfs['PESO'].sum():,.0f}" if not nfs.empty else "0",
                        "Valor Total": f"R$ {nfs['VALOR_NOTA'].sum():,.0f}" if not nfs.empty else "R$ 0",
                        "Status": status,
                    })

                if audit_rows:
                    st.dataframe(pd.DataFrame(audit_rows), use_container_width=True, hide_index=True)

    # Search in OBS
    st.markdown("---")
    st.markdown("### 🔎 Busca em Observações")
    search_obs = st.text_input(
        "Buscar nas observações de todas as viagens:",
        placeholder="Ex: b.o, azeplast, sabará...",
        key="exp_audit_obs_search",
    )

    if search_obs:
        all_viagens = read_viagens()
        if not all_viagens.empty and "OBS" in all_viagens.columns:
            obs_mask = all_viagens["OBS"].fillna("").str.contains(search_obs, case=False, na=False)
            results = all_viagens[obs_mask]
            if results.empty:
                st.info(f"Nenhuma viagem com '{search_obs}' nas observações.")
            else:
                st.success(f"🔍 {len(results)} viagem(ns) encontrada(s)")
                st.dataframe(
                    results[["DATA", "MOTORISTA", "PLACA", "OBS"]],
                    use_container_width=True,
                    hide_index=True,
                )


# ══════════════════════════════════════════════════════════════
# TAB 5 — CONFERÊNCIA (Expedição vs. Viagem)
# ══════════════════════════════════════════════════════════════
with tab_check:
    st.markdown("### ✅ Conferência — Expedição Planejada vs. Viagem Relatada")
    st.caption("Cruza automaticamente o plano de expedição com o relato do motorista (WhatsApp).")

    # Date selector
    exp_dates = get_expedition_dates()
    if not exp_dates:
        st.info("Nenhuma expedição salva ainda. Monte um plano na aba 📋 Planejar.")
    else:
        check_date = st.selectbox(
            "📅 Data da Expedição:",
            exp_dates,
            format_func=lambda d: d.strftime("%d/%m/%Y"),
            key="check_date",
        )

        st.markdown("---")

        # Check each vehicle
        any_result = False
        for placa in FLEET_PLATES:
            result = cross_check_expedition_vs_trip(check_date, placa)
            if result is None:
                continue

            any_result = True
            color = "#22d3ee" if result.get("viagem_registrada") else "#64748b"

            st.markdown(f"""
            <div style="
                background: rgba({_hex_to_rgb(color)}, 0.06);
                border: 1px solid rgba({_hex_to_rgb(color)}, 0.20);
                border-radius: 12px;
                padding: 16px 20px;
                margin-bottom: 16px;
            ">
                <div style="font-weight: 700; color: {color}; font-size: 16px; margin-bottom: 8px;">
                    🚚 {placa} — {result['data']}
                    {f"— {result.get('motorista', '')}" if result.get('motorista') else ""}
                </div>
            """, unsafe_allow_html=True)

            if result.get("viagem_registrada"):
                col_labels = ["", "Planejado", "Relatado (Wpp)"]
                rows_data = [
                    ["Pedidos", str(result["pedidos_planejados"]),
                     f"{result['pedidos_relatados']}  {result['status_pedidos']}"],
                    ["Entregas", "—",
                     f"{result['entregas_relatadas']}  {result['status_entregas']}"],
                    ["KM Rodado", "—", f"{result['km_rodado']} km"],
                ]
                if result.get("ocorrencias"):
                    rows_data.append(["Ocorrências", "—", result["ocorrencias"]])

                df_conf = pd.DataFrame(rows_data, columns=col_labels)
                st.dataframe(df_conf, use_container_width=True, hide_index=True)

                # If everything matches, allow marking as ENTREGUE
                if result["status_pedidos"] == "✅" and result["status_entregas"] == "✅":
                    st.markdown("""
                    <div style="color: #10b981; font-weight: 600; font-size: 13px; margin: 4px 0;">
                        ✅ Todos os pedidos foram entregues com sucesso!
                    </div>
                    """, unsafe_allow_html=True)

                    if st.button(
                        f"🏁 Marcar pedidos do {placa} como ENTREGUE",
                        key=f"deliver_{placa}",
                        use_container_width=True,
                        disabled=not is_adm(),
                    ):
                        n = mark_as_delivered(check_date, placa)
                        if n > 0:
                            st.success(f"✅ {n} pedido(s) marcados como ENTREGUE na Esteira!")
                        else:
                            st.warning("Nenhum pedido atualizado.")

                elif result["status_pedidos"] == "⚠️" or result["status_entregas"] == "⚠️":
                    st.markdown("""
                    <div style="color: #f59e0b; font-weight: 600; font-size: 13px; margin: 4px 0;">
                        ⚠️ Divergência detectada — revise antes de confirmar entrega.
                    </div>
                    """, unsafe_allow_html=True)
            else:
                st.markdown("""
                <div style="color: #64748b; font-size: 13px; padding: 8px 0;">
                    ⏳ Viagem ainda não registrada para este veículo/data.
                </div>
                """, unsafe_allow_html=True)

            st.markdown("</div>", unsafe_allow_html=True)

        if not any_result:
            st.info(f"Nenhum veículo atribuído na expedição de {check_date.strftime('%d/%m/%Y')}.")


# ══════════════════════════════════════════════════════════════
# TAB 6 — CAMINHO REVERSO (Confirmação Retroativa)
# ══════════════════════════════════════════════════════════════
with tab_reverse:
    st.markdown("### 🔄 Confirmação Reversa de Entregas")
    st.markdown("""
    <p style="color: #94a3b8; font-size: 13px; margin-bottom: 16px;">
        Confirme entregas que saíram sem checklist formal.
        Associe a <strong>Nota Fiscal</strong> ao pedido e verifique no <strong>Relatório de Entregas</strong>.
    </p>
    """, unsafe_allow_html=True)

    # Mode selector
    reverse_mode = st.radio(
        "Modo de entrada",
        ["📎 Upload de DANFE(s)", "✏️ Entrada Manual"],
        horizontal=True,
        key="reverse_mode",
        label_visibility="collapsed",
    )

    st.markdown("---")

    # ──────────────────────────────────────────────────────────
    # MODE 1: UPLOAD DANFE
    # ──────────────────────────────────────────────────────────
    if reverse_mode == "📎 Upload de DANFE(s)":
        st.markdown("#### 📎 Upload de DANFE(s) em PDF")
        st.caption("Faça upload de uma ou mais DANFEs. O sistema extrairá automaticamente NF, Pedido e Remetente.")

        danfe_files = st.file_uploader(
            "Selecione os PDFs de DANFE",
            type=["pdf"],
            accept_multiple_files=True,
            key="reverse_danfe_upload",
            label_visibility="collapsed",
            disabled=not is_adm(),
        )

        if danfe_files:
            with st.spinner("Processando DANFEs..."):
                extractions = batch_process_danfes(danfe_files)

            if not extractions:
                st.warning("Nenhum dado extraído dos PDFs.")
            else:
                st.success(f"📄 {len(extractions)} DANFE(s) processada(s)")

                # Build editable DataFrame
                edit_data = []
                for i, ext in enumerate(extractions):
                    edit_data.append({
                        "Arquivo": ext["arquivo"],
                        "NF": ext["nf"],
                        "Pedido": ext["pedido"],
                        "Remetente": ext["remetente"],
                        "Cliente": ext["cliente"],
                        "Peso": ext["peso"],
                        "Volumes": ext["volumes"],
                        "Status": "✅ OK" if ext["status"] == "extraido" and ext["nf"] and ext["pedido"] else "⚠️ Revisar",
                    })

                df_ext = pd.DataFrame(edit_data)

                # Editable table
                edited_ext = st.data_editor(
                    df_ext,
                    use_container_width=True,
                    hide_index=True,
                    num_rows="fixed",
                    key="reverse_danfe_editor",
                    column_config={
                        "Arquivo": st.column_config.TextColumn("📄 Arquivo", width="medium", disabled=True),
                        "NF": st.column_config.TextColumn("📋 NF", width="small"),
                        "Pedido": st.column_config.TextColumn("📦 Pedido", width="small"),
                        "Remetente": st.column_config.SelectboxColumn(
                            "🏢 Remetente",
                            options=EMPRESA_OPTIONS,
                            width="medium",
                        ),
                        "Cliente": st.column_config.TextColumn("👤 Cliente", width="medium", disabled=True),
                        "Peso": st.column_config.TextColumn("⚖️ Peso", width="small", disabled=True),
                        "Volumes": st.column_config.TextColumn("📦 Vol", width="small", disabled=True),
                        "Status": st.column_config.TextColumn("📌 Status", width="small", disabled=True),
                    },
                )

                # Filter valid rows (have both NF and Pedido)
                valid_mask = (
                    edited_ext["NF"].astype(str).str.strip().ne("")
                    & edited_ext["Pedido"].astype(str).str.strip().ne("")
                    & edited_ext["Remetente"].astype(str).str.strip().ne("")
                )
                valid_count = valid_mask.sum()

                st.markdown(f"""
                <div style="background: rgba(99, 102, 241, 0.05); border: 1px solid rgba(99, 102, 241, 0.15);
                    border-radius: 8px; padding: 12px; margin: 8px 0;">
                    <span style="color: #818cf8; font-weight: 600;">
                        {valid_count} de {len(edited_ext)} DANFE(s) prontas para confirmação
                    </span>
                </div>
                """, unsafe_allow_html=True)

                col_verify, col_confirm = st.columns(2)

                with col_verify:
                    if st.button(
                        "🔍 Verificar no Relatório",
                        type="secondary",
                        use_container_width=True,
                        disabled=(valid_count == 0) or not is_adm(),
                        key="reverse_verify_batch",
                    ):
                        verification_results = []
                        for _, row in edited_ext[valid_mask].iterrows():
                            nf = str(row["NF"]).strip()
                            remetente = str(row["Remetente"]).strip()
                            delivery = verify_delivery_in_report(nf, remetente)
                            verification_results.append({
                                "Pedido": row["Pedido"],
                                "NF": nf,
                                "Remetente": remetente,
                                "No Relatório": "✅ Encontrada" if delivery else "❌ Não encontrada",
                                "_found": delivery is not None,
                                "Veículo": delivery.get("VEICULO", "") if delivery else "",
                                "Data Entrega": (
                                    delivery["DATA"].strftime("%d/%m/%Y")
                                    if delivery and hasattr(delivery.get("DATA"), "strftime")
                                    else ""
                                ),
                            })

                        st.session_state["reverse_verification"] = verification_results
                        st.rerun()

                with col_confirm:
                    if st.button(
                        "✅ Confirmar Todas",
                        type="primary",
                        use_container_width=True,
                        disabled=(valid_count == 0) or not is_adm(),
                        key="reverse_confirm_batch",
                        help="Confirma TODAS as DANFEs — encontradas e não encontradas",
                    ):
                        with st.spinner(f"Processando {valid_count} pedidos em lote..."):
                            batch_items = [
                                {
                                    "pedido": str(row["Pedido"]).strip(),
                                    "nf": str(row["NF"]).strip(),
                                    "remetente": str(row["Remetente"]).strip(),
                                }
                                for _, row in edited_ext[valid_mask].iterrows()
                            ]
                            results = confirm_delivery_batch(batch_items)

                        st.session_state["reverse_results"] = results
                        st.rerun()

                # Show verification results
                if "reverse_verification" in st.session_state:
                    st.markdown("#### 🔍 Resultado da Verificação")
                    vdf = pd.DataFrame(st.session_state["reverse_verification"])

                    # Contadores
                    found_count = sum(1 for v in st.session_state["reverse_verification"] if v["_found"])
                    not_found_count = len(st.session_state["reverse_verification"]) - found_count

                    vc1, vc2 = st.columns(2)
                    with vc1:
                        st.metric("✅ Encontradas", found_count)
                    with vc2:
                        st.metric("❌ Não encontradas", not_found_count)

                    # Exibe tabela sem a coluna interna _found
                    display_cols = [c for c in vdf.columns if not c.startswith("_")]
                    st.dataframe(vdf[display_cols], use_container_width=True, hide_index=True)

                    # ── BOTÃO: Confirmar APENAS Encontradas ──
                    if found_count > 0:
                        st.markdown(f"""
                        <div style="background: rgba(34, 197, 94, 0.08); border: 1px solid rgba(34, 197, 94, 0.25);
                            border-radius: 8px; padding: 12px; margin: 8px 0;">
                            <span style="color: #22c55e; font-weight: 600;">
                                🎯 {found_count} pedido(s) encontrado(s) no relatório — prontos para ENTREGUE
                            </span>
                            <br><span style="color: #94a3b8; font-size: 0.85em;">
                                Os {not_found_count} pedido(s) não encontrados terão apenas a NF registrada, sem alterar o status.
                            </span>
                        </div>
                        """, unsafe_allow_html=True)

                        if st.button(
                            f"✅ Confirmar Apenas Encontradas ({found_count})",
                            type="primary",
                            use_container_width=True,
                            key="reverse_confirm_found_only",
                            disabled=not is_adm(),
                        ):
                            found_items = [
                                v for v in st.session_state["reverse_verification"] if v["_found"]
                            ]

                            with st.spinner(f"Processando {len(found_items)} pedidos encontrados em lote..."):
                                batch_items = [
                                    {
                                        "pedido": str(item["Pedido"]).strip(),
                                        "nf": str(item["NF"]).strip(),
                                        "remetente": str(item["Remetente"]).strip(),
                                    }
                                    for item in found_items
                                ]
                                results = confirm_delivery_batch(batch_items)

                            # Adiciona os não encontrados como info (NF registrada, sem alterar status)
                            not_found_items = [
                                v for v in st.session_state["reverse_verification"] if not v["_found"]
                            ]
                            for item in not_found_items:
                                results.append({
                                    "success": True,
                                    "pedido": str(item["Pedido"]).strip(),
                                    "nf": str(item["NF"]).strip(),
                                    "found_in_controle": True,
                                    "nf_conflict": False,
                                    "nf_saved": False,
                                    "found_in_report": False,
                                    "delivery_data": None,
                                    "status_updated": False,
                                    "location": None,
                                    "status_anterior": None,
                                    "message": (
                                        f"⏭️ Pedido {item['Pedido']}: NF {item['NF']} não encontrada "
                                        f"no relatório — status mantido (ignorado)."
                                    ),
                                })

                            st.session_state["reverse_results"] = results
                            st.session_state.pop("reverse_verification", None)
                            st.rerun()

                # Show confirmation results
                if "reverse_results" in st.session_state:
                    st.markdown("#### 📊 Resultado da Confirmação")
                    results = st.session_state["reverse_results"]
                    entregue_count = sum(1 for r in results if r.get("status_updated"))
                    nf_only_count = sum(1 for r in results if r["success"] and not r.get("status_updated") and not r.get("nf_conflict"))
                    conflict_count = sum(1 for r in results if r.get("nf_conflict"))
                    fail_count = sum(1 for r in results if not r["success"] and not r.get("nf_conflict"))

                    rc1, rc2, rc3 = st.columns(3)
                    with rc1:
                        st.metric("✅ ENTREGUE", entregue_count)
                    with rc2:
                        st.metric("📝 NF registrada", nf_only_count)
                    with rc3:
                        st.metric("⚠️ Conflito/Erro", conflict_count + fail_count)

                    for r in results:
                        if r.get("nf_conflict"):
                            st.warning(f"⚠️ {r['message']}")
                        elif r["success"] and r.get("status_updated"):
                            st.success(f"{r['message']}")
                        elif r["success"]:
                            st.info(f"{r['message']}")
                        else:
                            st.error(f"❌ {r['message']}")

                    if st.button("🗑️ Limpar resultados", key="reverse_clear_results", disabled=not is_adm()):
                        st.session_state.pop("reverse_results", None)
                        st.session_state.pop("reverse_verification", None)
                        st.rerun()

    # ──────────────────────────────────────────────────────────
    # MODE 2: ENTRADA MANUAL
    # ──────────────────────────────────────────────────────────
    else:
        st.markdown("#### ✏️ Confirmação Manual")
        st.caption("Selecione um pedido pendente, informe a NF e o remetente.")

        # Load pending orders
        df_pending = get_pending_orders_without_nf()

        if df_pending.empty:
            st.info("🎉 Nenhum pedido pendente sem NF encontrado!")
        else:
            st.markdown(f"""
            <div style="background: rgba(251, 191, 36, 0.05); border: 1px solid rgba(251, 191, 36, 0.15);
                border-radius: 8px; padding: 12px; margin-bottom: 16px;">
                <span style="color: #fbbf24; font-weight: 600;">
                    📋 {len(df_pending)} pedido(s) sem NF
                </span>
            </div>
            """, unsafe_allow_html=True)

            # Form
            with st.form("reverse_manual_form", clear_on_submit=True):
                mc1, mc2, mc3 = st.columns(3)

                with mc1:
                    pedido_options = df_pending["PEDIDO"].astype(str).tolist()
                    selected_pedido = st.selectbox(
                        "Pedido",
                        pedido_options,
                        format_func=lambda x: (
                            f"#{x} — {df_pending[df_pending['PEDIDO'].astype(str) == x]['CLIENTE'].values[0]}"
                            if len(df_pending[df_pending['PEDIDO'].astype(str) == x]) > 0
                            else f"#{x}"
                        ),
                    )

                with mc2:
                    manual_nf = st.text_input(
                        "Número da NF",
                        placeholder="Ex: 2426",
                    )

                with mc3:
                    manual_remetente = st.selectbox(
                        "Remetente",
                        EMPRESA_OPTIONS,
                    )

                submitted = st.form_submit_button(
                    "🔍 Verificar e Confirmar",
                    type="primary",
                    use_container_width=True,
                    disabled=not is_adm()
                )

                if submitted:
                    if not manual_nf.strip():
                        st.error("Informe o número da NF.")
                    else:
                        # Verify in report
                        delivery = verify_delivery_in_report(manual_nf.strip(), manual_remetente)

                        if delivery:
                            st.success(
                                f"✅ NF {manual_nf} encontrada no relatório! "
                                f"Veículo: {delivery.get('VEICULO', 'N/A')} | "
                                f"Cliente: {delivery.get('CLIENTE', 'N/A')}"
                            )
                        else:
                            st.warning(
                                f"⚠️ NF {manual_nf} NÃO encontrada no relatório de entregas "
                                f"(Remetente: {manual_remetente}). "
                                f"A confirmação será manual."
                            )

                        # Confirm delivery
                        result = confirm_delivery(selected_pedido, manual_nf.strip(), manual_remetente)

                        if result["success"]:
                            st.success(f"🎉 {result['message']}")
                            if result.get("status_anterior"):
                                st.info(f"Status anterior: {result['status_anterior']} → ENTREGUE")
                            st.balloons()
                        else:
                            st.error(f"❌ {result['message']}")

            # Show pending orders table
            st.markdown("---")
            st.markdown("##### 📋 Pedidos pendentes (sem NF)")
            display_pending = df_pending[["PEDIDO", "CLIENTE", "EMPRESA", "STATUS", "DATA"]].copy()
            if "DATA" in display_pending.columns:
                display_pending["DATA"] = pd.to_datetime(display_pending["DATA"], errors="coerce").dt.strftime("%d/%m/%Y")
            st.dataframe(
                display_pending,
                use_container_width=True,
                hide_index=True,
                column_config={
                    "PEDIDO": st.column_config.TextColumn("📋 Pedido", width="small"),
                    "CLIENTE": st.column_config.TextColumn("👤 Cliente", width="medium"),
                    "EMPRESA": st.column_config.TextColumn("🏢 Empresa", width="small"),
                    "STATUS": st.column_config.TextColumn("🔖 Status", width="small"),
                    "DATA": st.column_config.TextColumn("📅 Data", width="small"),
                },
            )

