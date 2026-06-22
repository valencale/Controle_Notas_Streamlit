"""
3_📄_Ingestao_PDF.py — Página de ingestão e processamento de PDFs.

Funcionalidades:
- Upload de PDFs (Notas/Romaneios Onfinity)
- Extração automática via regex (Pedido, Cliente, Endereço, Obs)
- Preview e edição manual dos dados extraídos
- Inserção na esteira com um clique
"""

import streamlit as st
import pandas as pd
from datetime import datetime
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from modules.ui_components import inject_custom_css, render_header
from modules.pdf_parser import extract_data_from_pdf, extract_text_preview
from modules.excel_handler import insert_pedido
from config import STATUS_OPTIONS, EMPRESA_OPTIONS
from modules.auth import is_adm

# ══════════════════════════════════════════════════════════════
# PAGE CONFIG
# ══════════════════════════════════════════════════════════════
st.set_page_config(page_title="Ingestão PDF — Gestão Logística", page_icon="📄", layout="wide")
inject_custom_css()

# ══════════════════════════════════════════════════════════════
# HEADER
# ══════════════════════════════════════════════════════════════
render_header("Ingestão de PDFs", "Importar Notas e Romaneios automaticamente")

st.markdown("---")

# ══════════════════════════════════════════════════════════════
# FILE UPLOADER
# ══════════════════════════════════════════════════════════════
st.markdown("""
<div style="text-align: center; padding: 16px 0;">
    <p style="color: #94a3b8; font-size: 14px;">
        📎 Faça upload de PDFs do tipo <strong>Onfinity — Mapa de Separação por Pedido</strong>
    </p>
</div>
""", unsafe_allow_html=True)

uploaded_files = st.file_uploader(
    "Upload de PDFs",
    type=["pdf"],
    accept_multiple_files=True,
    label_visibility="collapsed",
    help="Aceita múltiplos PDFs. Formato suportado: Onfinity 'Mapa de Separação por Pedido'",
    disabled=not is_adm(),
)

if not uploaded_files:
    st.markdown("""
    <div style="text-align: center; padding: 60px 0;">
        <div style="font-size: 72px; margin-bottom: 16px;">📄</div>
        <h3 style="color: #94a3b8; font-weight: 600;">Nenhum PDF selecionado</h3>
        <p style="color: #64748b; font-size: 14px;">
            Arraste e solte ficheiros PDF ou clique em "Browse files" acima
        </p>
    </div>
    """, unsafe_allow_html=True)
    st.stop()

# ══════════════════════════════════════════════════════════════
# PROCESS EACH PDF
# ══════════════════════════════════════════════════════════════
st.markdown("---")
st.markdown(f"### 📋 {len(uploaded_files)} PDF(s) carregado(s)")
st.info("💡 Cada página do PDF será analisada individualmente para identificar pedidos únicos.")

all_extracted = []

for file_idx, uploaded_file in enumerate(uploaded_files):
    with st.expander(f"📄 {uploaded_file.name}", expanded=True):

        # Extract data
        with st.status(f"Processando {uploaded_file.name}...", expanded=False) as status:
            try:
                extracted_list = extract_data_from_pdf(uploaded_file)
                uploaded_file.seek(0)  # Reset file pointer for potential re-read
                status.update(label=f"✅ {uploaded_file.name} processado com sucesso!", state="complete", expanded=False)
                if extracted_list:
                    st.toast(f"✅ {len(extracted_list)} pedido(s) extraído(s) de {uploaded_file.name}", icon="📄")
            except Exception as e:
                status.update(label=f"❌ Erro ao processar {uploaded_file.name}", state="error")
                st.error(f"Detalhes do erro crítico: {str(e)}")
                st.toast(f"❌ Falha ao processar {uploaded_file.name}", icon="🚨")
                continue

        if not extracted_list:
            st.warning(f"⚠️ Nenhum dado útil pôde ser extraído de {uploaded_file.name}. Verifique se o formato está correto.")

            # Show raw text for debugging
            with st.popover("🔍 Ver texto extraído (debug)"):
                try:
                    uploaded_file.seek(0)
                    raw_text = extract_text_preview(uploaded_file)
                    uploaded_file.seek(0)
                    st.code(raw_text, language="text")
                except Exception as e:
                    st.error(f"Erro ao ler texto: {e}")
            continue

        for item_idx, extracted in enumerate(extracted_list):
            cont_pages = extracted.get("_CONTINUATION_PAGES", 0)

            # Badge de continuação para pedidos multi-página
            if cont_pages > 0:
                st.markdown(f"""
                <div style="background: rgba(251, 191, 36, 0.08); border: 1px solid rgba(251, 191, 36, 0.25); border-radius: 12px; padding: 16px; margin-bottom: 12px;">
                    <p style="color: #fbbf24; font-weight: 700; font-size: 14px; margin-bottom: 4px;">
                        ✨ Dados extraídos automaticamente
                        <span style="background: rgba(251, 191, 36, 0.15); color: #f59e0b; padding: 2px 10px; border-radius: 20px; font-size: 12px; margin-left: 8px;">
                            📄 {cont_pages + 1} págs ({cont_pages} continuação{"ões" if cont_pages > 1 else ""} mesclada{"s" if cont_pages > 1 else ""})
                        </span>
                    </p>
                </div>
                """, unsafe_allow_html=True)
            else:
                st.markdown("""
                <div style="background: rgba(99, 102, 241, 0.05); border: 1px solid rgba(99, 102, 241, 0.15); border-radius: 12px; padding: 16px; margin-bottom: 12px;">
                    <p style="color: #818cf8; font-weight: 700; font-size: 14px; margin-bottom: 8px;">
                        ✨ Dados extraídos automaticamente
                    </p>
                </div>
                """, unsafe_allow_html=True)

            # Editable form for extracted data
            key_prefix = f"pdf_{file_idx}_{item_idx}"

            col1, col2 = st.columns(2)
            with col1:
                edit_pedido = st.text_input(
                    "Número do Pedido",
                    value=extracted.get("PEDIDO", ""),
                    key=f"{key_prefix}_pedido",
                )
            with col2:
                edit_id_cliente = st.text_input(
                    "ID Cliente",
                    value=extracted.get("ID_CLIENTE", ""),
                    key=f"{key_prefix}_id_cliente",
                )

            col3, col4, col5 = st.columns(3)
            with col3:
                edit_cliente = st.text_input(
                    "Cliente",
                    value=extracted.get("CLIENTE", ""),
                    key=f"{key_prefix}_cliente",
                )
            with col4:
                edit_empresa = st.selectbox(
                    "Empresa",
                    EMPRESA_OPTIONS,
                    index=EMPRESA_OPTIONS.index(extracted.get("EMPRESA", EMPRESA_OPTIONS[0])) if extracted.get("EMPRESA") in EMPRESA_OPTIONS else 0,
                    key=f"{key_prefix}_empresa",
                )
            with col5:
                edit_status = st.selectbox(
                    "Status Inicial",
                    STATUS_OPTIONS,
                    index=0,
                    key=f"{key_prefix}_status",
                )

            edit_endereco = st.text_input(
                "📍 Endereço de Entrega",
                value=extracted.get("ENDERECO", ""),
                key=f"{key_prefix}_endereco",
            )

            edit_obs = st.text_area(
                "📝 Observações",
                value=extracted.get("OBS", ""),
                key=f"{key_prefix}_obs",
                height=68,
            )

            # Store the edited data
            all_extracted.append({
                "key_prefix": key_prefix,
                "file_name": uploaded_file.name,
                "PEDIDO": edit_pedido,
                "ID_CLIENTE": edit_id_cliente,
                "CLIENTE": edit_cliente,
                "EMPRESA": edit_empresa,
                "ENDERECO": edit_endereco,
                "STATUS": edit_status,
                "OBS": edit_obs,
                "DATA": datetime.now(),
            })

            # Show raw text toggle
            with st.popover("🔍 Ver texto bruto"):
                try:
                    uploaded_file.seek(0)
                    raw_text = extract_text_preview(uploaded_file)
                    uploaded_file.seek(0)
                    st.code(raw_text[:2000], language="text")
                except Exception as e:
                    st.error(f"Erro: {e}")

# ══════════════════════════════════════════════════════════════
# INSERT ALL BUTTON
# ══════════════════════════════════════════════════════════════
if all_extracted:
    st.markdown("---")

    # Summary table
    st.markdown(f"### 📊 Resumo — {len(all_extracted)} pedido(s) para inserir")

    summary_df = pd.DataFrame([{
        "Pedido": item["PEDIDO"],
        "ID": item.get("ID_CLIENTE", ""),
        "Cliente": item["CLIENTE"],
        "Empresa": item["EMPRESA"],
        "Status": item["STATUS"],
        "Endereço": item["ENDERECO"][:50] + "..." if len(item.get("ENDERECO", "")) > 50 else item.get("ENDERECO", ""),
    } for item in all_extracted])

    st.dataframe(summary_df, width="stretch", hide_index=True)

    # Insert button
    col_insert, col_space = st.columns([1, 2])
    with col_insert:
        if st.button("🚀 Inserir Todos na Esteira", type="primary", use_container_width=True, disabled=not is_adm()):
            success_count = 0
            error_count = 0

            # Enhanced progress tracking
            progress_text = "Processando inserções. Por favor, aguarde..."
            progress_bar = st.progress(0, text=progress_text)

            total_items = len(all_extracted)
            for i, item in enumerate(all_extracted):
                try:
                    insert_pedido({
                        "DATA": item["DATA"],
                        "CLIENTE": item["CLIENTE"].upper().strip(),
                        "PEDIDO": item["PEDIDO"].strip(),
                        "EMPRESA": item["EMPRESA"],
                        "STATUS": item["STATUS"],
                        "OBS": item["OBS"].strip(),
                        "ENDERECO": item["ENDERECO"].strip(),
                        "NF": "",
                    })
                    success_count += 1
                except ValueError as e:
                    st.toast(f"⚠️ Aviso no pedido #{item['PEDIDO']}: {e}", icon="⚠️")
                    error_count += 1
                except Exception as e:
                    st.toast(f"❌ Erro no pedido #{item['PEDIDO']}: {e}", icon="🚨")
                    error_count += 1

                progress_percent = int(((i + 1) / total_items) * 100)
                progress_bar.progress(progress_percent, text=f"Inserindo pedido {i+1} de {total_items}...")

            progress_bar.empty() # Clear the progress bar after completion

            if success_count > 0:
                st.success(f"✅ {success_count} pedido(s) inserido(s) com sucesso!")
                st.toast(f"Inserção concluída: {success_count} sucesso(s)", icon="🎉")
            if error_count > 0:
                st.warning(f"⚠️ {error_count} pedido(s) com erro (verifique os avisos).")

            if success_count > 0:
                st.balloons()

