"""
6_📄_Extracao_Notas.py — Módulo de Extração de Nota Fiscal (DANFE).

Processa PDFs de Notas Fiscais Eletrônicas (DANFE), extrai dados
estruturados usando pdfplumber + regex, e exporta DataFrame formatado
para Excel (.xlsx) com colunas na ordem estrita definida pelo negócio.
"""

import streamlit as st
import pandas as pd
import io
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from modules.ui_components import inject_custom_css, render_header
from modules.danfe_parser import (
    extrair_multiplos_danfe,
    gerar_excel_danfe,
    COLUNAS_SAIDA,
)
from modules.auth import is_adm

# ══════════════════════════════════════════════════════════════
# PAGE CONFIG
# ══════════════════════════════════════════════════════════════
st.set_page_config(page_title="Extração de Notas Fiscais", page_icon="📄", layout="wide")
inject_custom_css()

# CSS adicional para esta página
st.markdown("""
<style>
    .kpi-row { display: flex; gap: 12px; margin-bottom: 20px; flex-wrap: wrap; }
    .kpi-box {
        flex: 1; min-width: 140px; padding: 18px 20px; border-radius: 14px;
        background: linear-gradient(135deg, #1e293b, #1a1f2e);
        border: 1px solid rgba(99,102,241,0.15);
        text-align: center;
    }
    .kpi-box .kpi-val { font-size: 28px; font-weight: 700; color: #f8fafc; }
    .kpi-box .kpi-lab { font-size: 11px; color: #94a3b8; text-transform: uppercase;
        letter-spacing: 0.5px; margin-top: 4px; }
    .kpi-green  { border-left: 4px solid #22c55e; }
    .kpi-amber  { border-left: 4px solid #f59e0b; }
    .kpi-red    { border-left: 4px solid #ef4444; }
    .kpi-blue   { border-left: 4px solid #6366f1; }
    .kpi-purple { border-left: 4px solid #a855f7; }

    .extraction-log {
        background: #0f172a;
        border: 1px solid rgba(99,102,241,0.12);
        border-radius: 10px;
        padding: 16px 20px;
        margin-top: 12px;
        font-family: 'Cascadia Code', 'Fira Code', monospace;
        font-size: 13px;
        color: #94a3b8;
        max-height: 200px;
        overflow-y: auto;
    }
    .extraction-log .log-ok   { color: #22c55e; }
    .extraction-log .log-warn { color: #f59e0b; }
    .extraction-log .log-err  { color: #ef4444; }

    .highlight-row {
        background-color: rgba(239,68,68,0.15) !important;
        color: #fca5a5 !important;
    }
</style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════
# HEADER
# ══════════════════════════════════════════════════════════════
render_header(
    "Extração de Notas Fiscais",
    "Processamento de DANFE — PDF → Excel estruturado"
)
st.markdown("---")

# ══════════════════════════════════════════════════════════════
# UPLOAD DE ARQUIVOS
# ══════════════════════════════════════════════════════════════
st.markdown("### 📄 Upload de Notas Fiscais (DANFE)")
arquivos_nf = st.file_uploader(
    "Arraste ou selecione os PDFs das Notas Fiscais",
    type=["pdf"],
    accept_multiple_files=True,
    key="uploader_danfe",
    help="Aceita múltiplos arquivos PDF de DANFE (Nota Fiscal Eletrônica)",
    disabled=not is_adm(),
)

st.markdown("---")

# ══════════════════════════════════════════════════════════════
# BOTÃO DE EXTRAÇÃO
# ══════════════════════════════════════════════════════════════
btn_disabled = not arquivos_nf
btn = st.button(
    "🔍 Extrair Dados",
    use_container_width=True,
    type="primary",
    disabled=btn_disabled or not is_adm(),
)

if btn_disabled and not btn:
    st.markdown("""
    <div style="text-align:center; padding:50px 0;">
        <div style="font-size:64px; margin-bottom:12px;">📄</div>
        <h3 style="color:#94a3b8;">Anexe os PDFs das Notas Fiscais acima</h3>
        <p style="color:#64748b; font-size:14px;">
            Faça upload de um ou mais arquivos PDF de DANFE.<br>
            Os dados serão extraídos automaticamente e formatados para exportação Excel.
        </p>
        <div style="margin-top:24px; padding:16px; background:#1e293b; border-radius:12px;
                    border:1px solid rgba(99,102,241,0.12); display:inline-block; text-align:left;">
            <p style="color:#94a3b8; font-size:12px; font-weight:600; text-transform:uppercase;
                      letter-spacing:0.5px; margin-bottom:8px;">Campos Extraídos:</p>
            <p style="color:#64748b; font-size:13px; line-height:1.8; margin:0;">
                📋 <strong style="color:#e2e8f0;">Nota Fiscal</strong> — Número da NF-e<br>
                📦 <strong style="color:#e2e8f0;">Pedido</strong> — Número do pedido associado<br>
                📅 <strong style="color:#e2e8f0;">Data / Mês</strong> — Emissão formatada<br>
                🏢 <strong style="color:#e2e8f0;">Remetente / Cliente</strong> — Razão Social<br>
                ⚖️ <strong style="color:#e2e8f0;">Peso / Volumes</strong> — Mercadoria<br>
                💰 <strong style="color:#e2e8f0;">Valor da Nota</strong> — Total NF
            </p>
        </div>
    </div>
    """, unsafe_allow_html=True)
    st.stop()

if not btn:
    # Se já processou antes, mostra resultados do session_state
    if "danfe_result" not in st.session_state:
        st.stop()

# ══════════════════════════════════════════════════════════════
# PROCESSAMENTO
# ══════════════════════════════════════════════════════════════
if btn:
    with st.status("🔍 Processando DANFE...", expanded=True) as status:
        total = len(arquivos_nf)
        st.write(f"📄 {total} arquivo(s) carregado(s)")

        # Progresso
        progress_bar = st.progress(0, text="Iniciando extração...")
        log_entries = []

        # Processa cada arquivo individualmente para dar feedback granular
        registros = []
        erros = 0

        for i, arquivo in enumerate(arquivos_nf):
            nome = arquivo.name
            progress_bar.progress(
                (i + 1) / total,
                text=f"Processando {i+1}/{total}: {nome}"
            )

            from modules.danfe_parser import extrair_danfe
            try:
                resultado = extrair_danfe(arquivo)
            except Exception as e:
                resultado = None
                st.toast(f"❌ Erro crítico no arquivo {nome}: {str(e)}", icon="🚨")

            if resultado and resultado.get('Nota_Fiscal'):
                registros.append(resultado)
                log_entries.append(
                    f'<span class="log-ok">✅</span> {nome} → '
                    f'NF {resultado["Nota_Fiscal"]}'
                    f'{" | Ped " + resultado["Pedido"] if resultado.get("Pedido") else ""}'
                    f' | {resultado["Cliente"]} | '
                    f'R$ {resultado["Valor_Nota"]}'
                )
            elif resultado:
                registros.append(resultado)
                erros += 1
                log_entries.append(
                    f'<span class="log-warn">⚠️</span> {nome} → '
                    f'Extração parcial (campos faltantes)'
                )
            else:
                registros.append({col: "" for col in COLUNAS_SAIDA})
                erros += 1
                log_entries.append(
                    f'<span class="log-err">❌</span> {nome} → '
                    f'Falha na extração'
                )

        progress_bar.progress(1.0, text="Processamento concluído!")

        # Cria DataFrame
        df_resultado = pd.DataFrame(registros, columns=COLUNAS_SAIDA)
        st.session_state.danfe_result = df_resultado
        st.session_state.danfe_log = log_entries
        st.session_state.danfe_erros = erros

        if erros == 0:
            status.update(
                label=f"✅ {total} nota(s) extraída(s) com sucesso!",
                state="complete"
            )
        elif erros < total:
            status.update(
                label=f"⚠️ {total - erros}/{total} extraídas ({erros} com avisos)",
                state="complete"
            )
        else:
            status.update(
                label=f"❌ Falha na extração de todas as notas",
                state="error"
            )

        # Log de extração
        st.markdown('<div class="extraction-log">' +
                    '<br>'.join(log_entries) +
                    '</div>', unsafe_allow_html=True)

    st.toast(f"✅ {total - erros} nota(s) processada(s)!", icon="📄")

# ══════════════════════════════════════════════════════════════
# RESULTADOS
# ══════════════════════════════════════════════════════════════
if "danfe_result" not in st.session_state:
    st.stop()

df = st.session_state.danfe_result
log_entries = st.session_state.get("danfe_log", [])
erros = st.session_state.get("danfe_erros", 0)
total_notas = len(df)

# ── KPIs ─────────────────────────────────────────────────────
notas_ok = total_notas - erros

# Peso total (soma numérica)
def _peso_to_float(p):
    """Converte peso BR para float: '1547,50' → 1547.5"""
    if not p:
        return 0.0
    try:
        return float(str(p).replace('.', '').replace(',', '.'))
    except (ValueError, TypeError):
        return 0.0

peso_total = sum(_peso_to_float(p) for p in df['Peso'])
peso_fmt = f"{peso_total:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

# Valor total
def _valor_to_float(v):
    """Converte valor BR para float: '673,50' → 673.5"""
    if not v:
        return 0.0
    try:
        return float(str(v).replace('.', '').replace(',', '.'))
    except (ValueError, TypeError):
        return 0.0

valor_total = sum(_valor_to_float(v) for v in df['Valor_Nota'])
valor_fmt = f"R$ {valor_total:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

volumes_total = sum(int(v) if str(v).isdigit() else 0 for v in df['Volumes'])

st.markdown(f"""
<div class="kpi-row">
    <div class="kpi-box kpi-blue">
        <div class="kpi-val">{total_notas}</div>
        <div class="kpi-lab">Notas Processadas</div>
    </div>
    <div class="kpi-box kpi-green">
        <div class="kpi-val">{notas_ok}</div>
        <div class="kpi-lab">Extraídas com Sucesso</div>
    </div>
    <div class="kpi-box kpi-purple">
        <div class="kpi-val">{volumes_total}</div>
        <div class="kpi-lab">Total Volumes</div>
    </div>
    <div class="kpi-box kpi-amber">
        <div class="kpi-val">{peso_fmt} kg</div>
        <div class="kpi-lab">Peso Bruto Total</div>
    </div>
    <div class="kpi-box kpi-green">
        <div class="kpi-val">{valor_fmt}</div>
        <div class="kpi-lab">Valor Total NF</div>
    </div>
</div>
""", unsafe_allow_html=True)

# ── TABELA DE RESULTADOS ─────────────────────────────────────
st.subheader("📊 Dados Extraídos")
st.caption("Verifique os dados antes de exportar. Campos vazios indicam dados não encontrados no PDF.")

# Formatação condicional: destaca linhas com Nota_Fiscal vazia
def _highlight_incompleto(row):
    if not row["Nota_Fiscal"]:
        return ["background-color: rgba(239,68,68,0.15); color: #fca5a5;"] * len(row)
    return [""] * len(row)

styled_df = df.style.apply(_highlight_incompleto, axis=1)
st.dataframe(styled_df, use_container_width=True, hide_index=True, height=400)

# ── EXPORTAÇÃO ───────────────────────────────────────────────
st.markdown("---")
st.markdown("### 📥 Exportar Dados")

col_csv, col_xlsx = st.columns(2)

with col_csv:
    csv_data = df.to_csv(index=False, sep=";", encoding="utf-8-sig")
    st.download_button(
        "📄 Baixar CSV",
        csv_data.encode("utf-8-sig"),
        "notas_fiscais.csv",
        "text/csv",
        use_container_width=True,
    )

with col_xlsx:
    excel_bytes = gerar_excel_danfe(df)
    st.download_button(
        "📊 Baixar Excel (.xlsx)",
        excel_bytes,
        "notas_fiscais.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

# ── LOG DE EXTRAÇÃO (expansível) ─────────────────────────────
if log_entries:
    with st.expander("🔍 Log Detalhado de Extração", expanded=False):
        st.markdown(
            '<div class="extraction-log">' +
            '<br>'.join(log_entries) +
            '</div>',
            unsafe_allow_html=True
        )
