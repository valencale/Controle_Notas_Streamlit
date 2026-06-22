"""
9_📆_Historico_Expedicoes.py — Visão Histórica de Expedições (Mural).

Painel visual em formato de MURAL (quadrados lado a lado) dos dias que tiveram expedição.
Ao clicar em um dia, abre os detalhes dos veículos e pedidos.
"""

import streamlit as st
import pandas as pd
from datetime import datetime, timedelta

import sys
import os

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from modules.ui_components import inject_custom_css, render_header
from modules.auth import is_adm
from modules.expedition_engine import read_expeditions, _format_date_br, _parse_date_br
from config import FLEET_PLATES

# ══════════════════════════════════════════════════════════════
# PAGE CONFIG
# ══════════════════════════════════════════════════════════════
st.set_page_config(
    page_title="Histórico de Expedições",
    page_icon="📆",
    layout="wide",
)
inject_custom_css()

# ══════════════════════════════════════════════════════════════
# CUSTOM CSS MURAL & DETAILS
# ══════════════════════════════════════════════════════════════
st.markdown("""
<style>
    /* ═══ MURAL (QUADRADOS LADO A LADO) ═══ */
    .mural-card {
        background: linear-gradient(135deg, #1e293b 0%, #1a1f2e 100%);
        border: 1px solid rgba(99, 102, 241, 0.15);
        border-radius: 16px;
        padding: 20px;
        text-align: center;
        transition: all 0.2s ease;
        height: 100%;
        display: flex;
        flex-direction: column;
        justify-content: space-between;
    }
    .mural-card:hover {
        border-color: rgba(99, 102, 241, 0.4);
        box-shadow: 0 8px 24px rgba(99, 102, 241, 0.15);
        transform: translateY(-3px);
    }
    .mural-card.selected {
        border: 2px solid #6366f1;
        background: linear-gradient(135deg, #25334d 0%, #1e293b 100%);
        box-shadow: 0 0 20px rgba(99, 102, 241, 0.25);
    }
    .m-date {
        font-size: 20px; font-weight: 700; color: #f8fafc;
        margin-bottom: 4px;
    }
    .m-weekday {
        font-size: 11px; font-weight: 600; color: #818cf8;
        text-transform: uppercase; letter-spacing: 1px;
    }
    .m-kpis {
        display: flex; justify-content: space-around;
        margin: 16px 0; padding: 12px 0;
        background: rgba(15, 23, 42, 0.4);
        border-radius: 12px; border: 1px solid rgba(255, 255, 255, 0.05);
    }
    .m-kpi-col { display: flex; flex-direction: column; }
    .m-kpi-val { font-size: 18px; font-weight: 700; color: #f8fafc; }
    .m-kpi-lbl { font-size: 10px; color: #94a3b8; text-transform: uppercase; }

    /* ═══ DETALHES ABAIXO ═══ */
    .details-panel {
        background: rgba(15, 23, 42, 0.6);
        border: 1px solid rgba(99, 102, 241, 0.15);
        border-radius: 16px;
        padding: 30px;
        margin-top: 10px;
        animation: fadeIn 0.3s ease-out;
    }
    .vehicle-section {
        background: #1e293b;
        border-radius: 12px; border: 1px solid rgba(255,255,255,0.05);
        margin-bottom: 20px; overflow: hidden;
    }
    .vehicle-header {
        background: rgba(0, 0, 0, 0.2);
        padding: 14px 20px;
        display: flex; align-items: center; justify-content: space-between;
        border-bottom: 1px solid rgba(255,255,255,0.05);
    }
    .vehicle-plate {
        background: linear-gradient(135deg, #4f46e5, #6366f1);
        color: #fff; font-weight: 700; font-size: 13px;
        padding: 6px 14px; border-radius: 20px;
    }
    .vehicle-stats { color: #cbd5e1; font-size: 13px; font-weight: 500; }

    /* ═══ TABELA DE PEDIDOS ═══ */
    .table-container { padding: 0; }
    .tr-header {
        display: grid;
        grid-template-columns: 80px 80px 2fr 1.5fr 1fr 100px 70px 80px 90px;
        padding: 10px 16px;
        font-size: 11px; font-weight: 700; color: #94a3b8;
        text-transform: uppercase; letter-spacing: 0.5px;
        border-bottom: 1px solid rgba(255,255,255,0.05);
    }
    .tr-row {
        display: grid;
        grid-template-columns: 80px 80px 2fr 1.5fr 1fr 100px 70px 80px 90px;
        padding: 12px 16px;
        align-items: center; gap: 8px;
        border-bottom: 1px solid rgba(255,255,255,0.02);
        transition: background 0.15s;
    }
    .tr-row:hover { background: rgba(99, 102, 241, 0.05); }
    .tr-row:last-child { border-bottom: none; }
    
    .cell-num { font-weight: 700; color: #818cf8; font-size: 14px; }
    .cell-nf { 
        background: rgba(34, 197, 94, 0.12); color: #22c55e;
        font-weight: 600; font-size: 12px;
        padding: 3px 6px; border-radius: 8px; text-align: center;
        max-width: 80px;
    }
    .cell-nf.empty { background: rgba(239, 68, 68, 0.1); color: #ef4444; }
    .cell-text { color: #e2e8f0; font-size: 13px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
    .cell-sub { color: #94a3b8; font-size: 12px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
    .cell-val { color: #f59e0b; font-size: 13px; font-weight: 600; text-align: right; }
    .cell-money { color: #10b981; font-size: 13px; font-weight: 600; text-align: right; }

    details.vehicle-section > summary {
        list-style: none;
        cursor: pointer;
    }
    details.vehicle-section > summary::-webkit-details-marker {
        display: none;
    }
    .vehicle-plate-container { display: flex; align-items: center; }
    .vehicle-plate-container::before {
        content: '▶';
        color: #94a3b8;
        font-size: 12px;
        margin-right: 10px;
        transition: transform 0.2s;
    }
    details[open] > summary .vehicle-plate-container::before {
        transform: rotate(90deg);
    }

    @keyframes fadeIn {
        from { opacity: 0; transform: translateY(-10px); }
        to { opacity: 1; transform: translateY(0); }
    }
</style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════
# HEADER E LOAD DATA
# ══════════════════════════════════════════════════════════════
render_header(
    "Histórico de Expedições",
    "Visão em mural com detalhamento por veículo e empresa."
)
st.markdown("---")

@st.cache_data(ttl=60)
def _load_master_data():
    """Carrega expedições e tenta cruzar EMPRESA com base de dados da Esteira."""
    df_exp = read_expeditions()
    if df_exp.empty:
        return df_exp

    # Tenta carregar base mestre para pegar a coluna EMPRESA e VALOR se houver
    try:
        from modules.excel_handler import read_principal, read_historico
        df_p = read_principal()
        df_h = read_historico()
        df_master = pd.concat([df_p, df_h], ignore_index=True)
        # Prepara chaves: PEDIDO
        df_master["PEDIDO"] = df_master["PEDIDO"].astype(str).str.strip()
        df_exp["PEDIDO"] = df_exp["PEDIDO"].astype(str).str.strip()
        
        # Mapeia empresa
        map_empresa = df_master.drop_duplicates("PEDIDO").set_index("PEDIDO")["EMPRESA"].to_dict()
        df_exp["_EMPRESA_MASTER"] = df_exp["PEDIDO"].map(map_empresa)
    except Exception as e:
        df_exp["_EMPRESA_MASTER"] = None

    # Normaliza datas
    df_exp["_DATE"] = df_exp["DATA_EXPEDICAO"].apply(
        lambda x: _parse_date_br(_format_date_br(x)) if pd.notna(x) else None
    )
    df_exp = df_exp.dropna(subset=["_DATE"])
    
    # Trata veículos NULL/Vazios
    df_exp["VEICULO"] = df_exp["VEICULO"].fillna("SEM VEÍCULO")
    df_exp["VEICULO"] = df_exp["VEICULO"].apply(lambda x: "SEM VEÍCULO" if str(x).strip() in ("", "nan") else str(x).strip().upper())
    
    # Trata valores vazios/numéricos
    df_exp["PESO"] = pd.to_numeric(df_exp["PESO"], errors="coerce").fillna(0)
    df_exp["VOLUMES"] = pd.to_numeric(df_exp["VOLUMES"], errors="coerce").fillna(0)
    
    # NOVO BLOCO para carregar entregas_cache.xlsx
    try:
        if os.path.exists("entregas_cache.xlsx"):
            df_cache = pd.read_excel("entregas_cache.xlsx", header=1)
            
            def safe_nf(x):
                if pd.isna(x): return ""
                try:
                    return str(int(float(x)))
                except (ValueError, TypeError):
                    return str(x).strip()

            col_rem = [c for c in df_cache.columns if "Remetente" in str(c)][0]
            col_nf = [c for c in df_cache.columns if "Nota_Fiscal" in str(c)][0]
            col_op = [c for c in df_cache.columns if "Opera" in str(c)][0]
            col_val = [c for c in df_cache.columns if "Valor" in str(c)][0]
            col_data = [c for c in df_cache.columns if "Data" in str(c)][0]
            col_veiculo = [c for c in df_cache.columns if "Ve" in str(c) and "culo" in str(c)][0]
            col_cliente = [c for c in df_cache.columns if "Cliente" in str(c)][0]
            col_vol = [c for c in df_cache.columns if "Volume" in str(c)][0]
            col_peso = [c for c in df_cache.columns if "Peso" in str(c)][0]
            col_bairro = [c for c in df_cache.columns if "Bairro" in str(c)][0]
            col_uf = [c for c in df_cache.columns if "UF" in str(c)][0]
            
            # Trata destino
            df_cache["_DESTINO_CACHE"] = df_cache[col_bairro].fillna("").astype(str) + " - " + df_cache[col_uf].fillna("").astype(str)
            df_cache["_DESTINO_CACHE"] = df_cache["_DESTINO_CACHE"].str.strip(" -")
            
            # Limpa chave no df_cache
            df_cache["_KEY_CACHE"] = df_cache[col_rem].astype(str).str.strip().str.upper() + \
                                     df_cache[col_nf].apply(safe_nf)
            
            # Prepara a mesma chave no df_exp
            emp_series = df_exp["_EMPRESA_MASTER"].fillna("").astype(str).str.strip().str.upper()
            nf_series = df_exp["NF"].apply(safe_nf)
            df_exp["_KEY_EXP"] = emp_series + nf_series
            
            # Cria dicionários de mapeamento
            map_op = df_cache.drop_duplicates("_KEY_CACHE").set_index("_KEY_CACHE")[col_op].to_dict()
            map_val = df_cache.drop_duplicates("_KEY_CACHE").set_index("_KEY_CACHE")[col_val].to_dict()
            map_dest = df_cache.drop_duplicates("_KEY_CACHE").set_index("_KEY_CACHE")["_DESTINO_CACHE"].to_dict()
            
            df_exp["OPERACAO"] = df_exp["_KEY_EXP"].map(map_op)
            df_exp["VALOR"] = df_exp["_KEY_EXP"].map(map_val)
            df_exp["DESTINO"] = df_exp["_KEY_EXP"].map(map_dest)
            
            # ADICIONAR REGISTROS EXCLUSIVOS DO CACHE (Que nao estao em expedicoes.xlsx)
            missing_in_exp = df_cache[~df_cache["_KEY_CACHE"].isin(df_exp["_KEY_EXP"])].copy()
            if not missing_in_exp.empty:
                df_extra = pd.DataFrame()
                df_extra["PEDIDO"] = "N/A"
                df_extra["DATA_EXPEDICAO"] = pd.to_datetime(missing_in_exp[col_data], errors="coerce").dt.strftime("%d/%m/%Y")
                df_extra["_DATE"] = pd.to_datetime(missing_in_exp[col_data], errors="coerce").dt.date
                df_extra["VEICULO"] = missing_in_exp[col_veiculo].fillna("SEM VEÍCULO").astype(str).str.strip().str.upper()
                df_extra["MOTORISTA"] = ""
                df_extra["STATUS_ENTREGA"] = ""
                df_extra["NF"] = missing_in_exp[col_nf].apply(safe_nf)
                df_extra["CLIENTE"] = missing_in_exp[col_cliente].astype(str)
                df_extra["_EMPRESA_MASTER"] = missing_in_exp[col_rem].astype(str).str.strip().str.upper()
                df_extra["VOLUMES"] = pd.to_numeric(missing_in_exp[col_vol], errors="coerce").fillna(0)
                df_extra["PESO"] = pd.to_numeric(missing_in_exp[col_peso], errors="coerce").fillna(0)
                df_extra["VALOR"] = pd.to_numeric(missing_in_exp[col_val], errors="coerce").fillna(0.0)
                df_extra["DESTINO"] = missing_in_exp["_DESTINO_CACHE"]
                df_extra["OPERACAO"] = missing_in_exp[col_op].astype(str)
                df_extra["ORDEM_ENTREGA"] = 99
                
                df_exp = pd.concat([df_exp, df_extra], ignore_index=True)
                
    except Exception as e:
        pass

    # Trata os valores ausentes de Valor e Operacao
    if "VALOR" not in df_exp.columns:
        df_exp["VALOR"] = 0.0
    else:
        df_exp["VALOR"] = df_exp["VALOR"].fillna(0.0)
        
    if "OPERACAO" not in df_exp.columns:
        df_exp["OPERACAO"] = "null"
    else:
        df_exp["OPERACAO"] = df_exp["OPERACAO"].fillna("null")
        
    return df_exp.sort_values("_DATE", ascending=False)


df_all = _load_master_data()

if df_all.empty:
    st.info("Nenhuma expedição registrada no sistema.")
    st.stop()

# ══════════════════════════════════════════════════════════════
# FILTROS E GLOBAL STATUS
# ══════════════════════════════════════════════════════════════
WEEKDAYS_PT = {
    0: "Segunda", 1: "Terça", 2: "Quarta", 3: "Quinta", 4: "Sexta", 5: "Sábado", 6: "Domingo"
}

c1, c2, c3 = st.columns([1.5, 1.5, 1])
with c1:
    period = st.selectbox("📅 Período", ["Últimos 15 dias", "Últimos 30 dias", "Último mês", "Tudo"], index=0)
    days_map = {"Últimos 15 dias": 15, "Últimos 30 dias": 30, "Último mês": 60, "Tudo": 9999}
    cutoff = datetime.now().date() - timedelta(days=days_map[period])
with c2:
    all_plates = sorted(df_all["VEICULO"].unique())
    veic_filter = st.selectbox("🚚 Veículo", ["Todos"] + all_plates)
with c3:
    st.markdown("<br>", unsafe_allow_html=True)
    busca = st.text_input("🔍 Busca", placeholder="Pedido, NF...")

df_filtered = df_all[df_all["_DATE"] >= cutoff].copy()
if veic_filter != "Todos":
    df_filtered = df_filtered[df_filtered["VEICULO"] == veic_filter]
if busca:
    b = busca.strip().upper()
    mask = (
        df_filtered["PEDIDO"].astype(str).str.contains(b) |
        df_filtered["NF"].astype(str).str.contains(b) |
        df_filtered["CLIENTE"].astype(str).str.upper().str.contains(b)
    )
    df_filtered = df_filtered[mask]

unique_dates = sorted(df_filtered["_DATE"].unique(), reverse=True)

if not unique_dates:
    st.warning("Nenhum dado encontrado com os filtros atuais.")
    st.stop()

st.markdown("### 🗓️ Mural de Expedições")

# Inicializa state para o dia selecionado
if "hist_selected_date" not in st.session_state:
    st.session_state["hist_selected_date"] = unique_dates[0]

# Valida se o date no state ainda existe na lista filtrada, senão pega o primeiro
if st.session_state["hist_selected_date"] not in unique_dates:
    st.session_state["hist_selected_date"] = unique_dates[0]

# ══════════════════════════════════════════════════════════════
# RENDER MURAL (Grid)
# ══════════════════════════════════════════════════════════════
COLS_PER_ROW = 5
for i in range(0, len(unique_dates), COLS_PER_ROW):
    row_dates = unique_dates[i:i+COLS_PER_ROW]
    cols = st.columns(COLS_PER_ROW)
    
    for idx, dt in enumerate(row_dates):
        df_d = df_filtered[df_filtered["_DATE"] == dt]
        n_ped = len(df_d)
        n_veic = df_d["VEICULO"].nunique()
        n_nfs = df_d["NF"].apply(lambda x: 1 if pd.notna(x) and str(x).strip() not in ("","nan") else 0).sum()
        
        is_selected = dt == st.session_state["hist_selected_date"]
        sel_class = "selected" if is_selected else ""
        
        with cols[idx]:
            # Botão invisível sobreposto ao card via Streamlit nativo não é possível facilmente com CSS puro, 
            # então renderizamos um card HTML visual e o st.button atua como toggle logo abaixo.
            card_html = f"""
            <div class="mural-card {sel_class}">
                <div>
                    <div class="m-date">{dt.strftime('%d/%m/%Y')}</div>
                    <div class="m-weekday">{WEEKDAYS_PT.get(dt.weekday(), '')}</div>
                </div>
                <div class="m-kpis">
                    <div class="m-kpi-col">
                        <span class="m-kpi-val">{n_ped}</span>
                        <span class="m-kpi-lbl">Peds</span>
                    </div>
                    <div class="m-kpi-col">
                        <span class="m-kpi-val" style="color:#22c55e;">{n_nfs}</span>
                        <span class="m-kpi-lbl">NFs</span>
                    </div>
                    <div class="m-kpi-col">
                        <span class="m-kpi-val" style="color:#f59e0b;">{n_veic}</span>
                        <span class="m-kpi-lbl">Veic</span>
                    </div>
                </div>
            </div>
            """
            st.markdown(card_html, unsafe_allow_html=True)
            # Botão para setar o state (fica logo abaixo do quadrado visual)
            if st.button("Ver Detalhes", key=f"btn_{dt.isoformat()}", use_container_width=True):
                st.session_state["hist_selected_date"] = dt
                st.rerun()

st.markdown("---")

# ══════════════════════════════════════════════════════════════
# RENDER DETAILS (Painel expandido do dia selecionado)
# ══════════════════════════════════════════════════════════════
sel_dt = st.session_state["hist_selected_date"]
df_day = df_filtered[df_filtered["_DATE"] == sel_dt].copy()

st.markdown(f"### 📋 Detalhes — {sel_dt.strftime('%d/%m/%Y')} ({WEEKDAYS_PT.get(sel_dt.weekday(), '')})")

if df_day.empty:
    st.info("Nenhum dado encontrado para o dia selecionado (pode estar filtrado).")
else:
    # --- ADICIONAR BOTÃO DE EXPORTAÇÃO AQUI ---
    try:
        from modules.expedition_engine import export_checklist_excel
        excel_bytes = export_checklist_excel(sel_dt, df_override=df_day)
        st.download_button(
            label="📥 Exportar Expedição do Dia",
            data=excel_bytes,
            file_name=f"expedicao_{sel_dt.strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )
    except Exception as e:
        st.error(f"Erro ao gerar Excel: {e}")
        
    st.markdown('<div class="details-panel">', unsafe_allow_html=True)
    
    # Veículos do dia ordenados alfabeticamente
    veiculos_do_dia = sorted(df_day["VEICULO"].unique())
    
    for placa in veiculos_do_dia:
        df_v = df_day[df_day["VEICULO"] == placa].copy()
        
        n_ped_v = len(df_v)
        peso_v = df_v["PESO"].sum()
        vol_v = df_v["VOLUMES"].sum()
        valor_v = df_v.get("VALOR", pd.Series([0])).sum()
        motorista = df_v.iloc[0].get("MOTORISTA", "")
        motorista = "" if pd.isna(motorista) or motorista.lower()=="nan" else f" — {motorista}"
        
        # Inicia a tabela do veículo
        html_str = f"""
<details class="vehicle-section">
<summary class="vehicle-header">
<div class="vehicle-plate-container">
<span class="vehicle-plate">🚚 {placa}</span>
<span style="color:#f8fafc; font-weight:600; margin-left:12px; font-size:14px;">{motorista}</span>
</div>
<div class="vehicle-stats">{n_ped_v} pedidos | {peso_v:,.2f} kg | {int(vol_v)} volumes | R$ {valor_v:,.2f}</div>
</summary>
<div class="table-container">
<div class="tr-header">
<div>Pedido</div>
<div>NF</div>
<div>Cliente</div>
<div>Destino</div>
<div>Empresa</div>
<div>Operação</div>
<div style="text-align: right;">Volume</div>
<div style="text-align: right;">Peso</div>
<div style="text-align: right;">Valor (R$)</div>
</div>
"""
        
        # Sort por ordem de entrega
        df_v["_ORD"] = pd.to_numeric(df_v["ORDEM_ENTREGA"], errors="coerce").fillna(99)
        df_v = df_v.sort_values("_ORD")
        
        for _, row in df_v.iterrows():
            ped = str(row.get("PEDIDO")).strip()
            if ped.lower() in ("nan", "none", "n/a", ""):
                ped = "INDISPONIVEL"
                
            nf = str(row.get("NF")).strip()
            nf = "" if nf.lower() == "nan" else nf
            nf_class = "" if nf else "empty"
            nf_txt = f"NF {nf}" if nf else "S/ NF"
            
            cli = str(row.get("CLIENTE")).strip()
            cli = "" if cli.lower() == "nan" else cli
            
            # Puxa dados extras
            emp = str(row.get("_EMPRESA_MASTER", ""))
            emp = "null" if emp.lower() in ("nan", "none", "") else emp
            
            op = str(row.get("OPERACAO", ""))
            op = "null" if op.lower() in ("nan", "none", "") else op
            
            dest = str(row.get("DESTINO", ""))
            dest = "" if dest.lower() in ("nan", "none") else dest

            vol = int(row.get("VOLUMES", 0))
            peso = float(row.get("PESO", 0.0))
            val = float(row.get("VALOR", 0.0))
            
            html_str += f"""
<div class="tr-row">
<div class="cell-num">#{ped}</div>
<div class="cell-nf {nf_class}">{nf_txt}</div>
<div class="cell-text" title="{cli}">{cli}</div>
<div class="cell-sub" title="{dest}">{dest}</div>
<div class="cell-sub">{emp}</div>
<div class="cell-sub">{op}</div>
<div class="cell-val">{vol}</div>
<div class="cell-val">{peso:,.2f} kg</div>
<div class="cell-money">{val:,.2f}</div>
</div>
"""
            
        html_str += "</div></details>"
        st.markdown(html_str, unsafe_allow_html=True)
        
    st.markdown('</div>', unsafe_allow_html=True)
