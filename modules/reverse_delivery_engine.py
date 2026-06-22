"""
reverse_delivery_engine.py — Motor de Confirmação Reversa de Entregas.

Fluxo:
    1. Usuário associa NF ao pedido (via DANFE PDF ou entrada manual)
    2. Sistema busca NF no RELATÓRIO DE ENTREGAS (por REMETENTE+NF)
    3. Se encontrada → confirma entrega
    4. Atualiza STATUS para ENTREGUE e NF na CONTROLE NOTAS
    5. Salva registro retroativo no expedicoes.xlsx
    6. Opcionalmente move da aba Dados para Historico
"""

import os
import sys
import logging
from typing import Optional
from datetime import datetime

import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from config import STATUS_OPTIONS

logger = logging.getLogger(__name__)


# ══════════════════════════════════════════════════════════════
# HELPERS
# ══════════════════════════════════════════════════════════════

def _normalize_nf(nf) -> str:
    """Normaliza número de NF: converte float→int→str, preserva valor exato."""
    if pd.isna(nf) or nf is None:
        return ""
    # Float do Excel (2589.0) → int (2589) → str ("2589")
    if isinstance(nf, float):
        try:
            return str(int(nf))
        except (ValueError, OverflowError):
            return ""
    s = str(nf).strip()
    # Se vier como string "2589.0", converte via float→int
    try:
        return str(int(float(s)))
    except (ValueError, OverflowError):
        return s


def _normalize_remetente(remetente: str) -> str:
    """Normaliza nome do remetente para comparação."""
    return str(remetente).strip().upper()


# ══════════════════════════════════════════════════════════════
# VERIFICAÇÃO NO RELATÓRIO DE ENTREGAS
# ══════════════════════════════════════════════════════════════

def verify_delivery_in_report(nf: str, remetente: str) -> Optional[dict]:
    """
    Busca NF no RELATÓRIO DE ENTREGAS usando chave composta REMETENTE+NF.

    NFs são únicas por remetente, por isso a busca sempre exige ambos.

    Args:
        nf: Número da nota fiscal
        remetente: Nome do remetente (ONFINITY, GREEN BAGS, etc.)

    Returns:
        Dict com dados da entrega ou None se não encontrada.
    """
    from modules.delivery_reader import read_deliveries_report

    nf_norm = _normalize_nf(nf)
    rem_norm = _normalize_remetente(remetente)

    if not nf_norm:
        return None

    try:
        df = read_deliveries_report()
    except FileNotFoundError:
        logger.warning("Relatório de entregas não disponível.")
        return None

    if df.empty:
        return None

    # Normaliza NF do relatório para string (pode ser int ou str)
    df["_NF_NORM"] = df["NOTA_FISCAL"].apply(
        lambda x: str(int(x)) if pd.notna(x) and x != 0 else ""
    )
    df["_REM_NORM"] = df["REMETENTE"].apply(_normalize_remetente)

    # Busca: REMETENTE + NF (chave composta — NFs são únicas por remetente)
    matches = pd.DataFrame()
    if rem_norm:
        mask = (df["_NF_NORM"] == nf_norm) & (df["_REM_NORM"] == rem_norm)
        matches = df[mask]

    if matches.empty:
        return None

    # Retorna o primeiro match
    row = matches.iloc[0]
    return {
        "DATA": row.get("DATA"),
        "VEICULO": row.get("VEICULO", ""),
        "OPERACAO": row.get("OPERACAO", ""),
        "REMETENTE": row.get("REMETENTE", ""),
        "CLIENTE": row.get("CLIENTE", ""),
        "BAIRRO": row.get("BAIRRO", ""),
        "UF": row.get("UF", ""),
        "NOTA_FISCAL": nf_norm,
        "PESO": row.get("PESO", 0),
        "VOLUMES": row.get("VOLUMES", 0),
        "VALOR_NOTA": row.get("VALOR_NOTA", 0),
        "FRETE": row.get("FRETE", 0),
    }


# ══════════════════════════════════════════════════════════════
# ASSOCIAÇÃO NF → PEDIDO
# ══════════════════════════════════════════════════════════════

def associate_nf_to_order(pedido: str, nf: str) -> bool:
    """
    Associa NF ao pedido na CONTROLE NOTAS (coluna NF).

    Args:
        pedido: Número do pedido
        nf: Número da nota fiscal

    Returns:
        True se atualizado com sucesso.
    """
    from modules.excel_handler import update_nf_batch

    nf_clean = _normalize_nf(nf)
    count = update_nf_batch({str(pedido).strip(): nf_clean})
    return count > 0


# ══════════════════════════════════════════════════════════════
# CONFIRMAÇÃO DE ENTREGA (FLUXO COMPLETO)
# ══════════════════════════════════════════════════════════════

def confirm_delivery(pedido: str, nf: str, remetente: str) -> dict:
    """
    Fluxo completo de confirmação reversa (conforme regra de negócio):

    1. Com o PEDIDO extraído do PDF, busca na CONTROLE NOTAS
       → Se NÃO encontrar: notifica "pedido não encontrado"
    2. Se encontrar, verifica se já possui NF anotada
       → Se NÃO tiver NF: ANOTA o número da NF do PDF
       → Se já tiver NF: verifica se confere com a do PDF
         → Se diverge: notifica conflito
    3. Verifica se a NF existe no RELATÓRIO DE ENTREGAS
       → Se SIM: muda STATUS para ENTREGUE + salva expedição
       → Se NÃO: notifica "NF não encontrada no relatório"

    Regra: se a NF existe no relatório, foi ENTREGUE.

    Args:
        pedido: Número do pedido (extraído do rodapé do PDF)
        nf: Número da nota fiscal (extraído do PDF)
        remetente: Nome do remetente (extraído do PDF)

    Returns:
        Dict com resultado da operação.
    """
    from modules.excel_handler import (
        read_principal,
        read_historico,
        update_nf_batch,
        update_status_batch,
    )

    nf_clean = _normalize_nf(nf)
    pedido_clean = str(pedido).strip()

    result = {
        "success": False,
        "pedido": pedido_clean,
        "nf": nf_clean,
        "found_in_controle": False,
        "nf_existing": None,       # NF que já estava anotada
        "nf_conflict": False,      # True se NF do PDF diverge da anotada
        "nf_saved": False,
        "found_in_report": False,
        "delivery_data": None,
        "status_updated": False,
        "location": None,          # "Dados" or "Historico"
        "status_anterior": None,
        "message": "",
    }

    # ══════════════════════════════════════════════════════════
    # PASSO 1: Buscar PEDIDO na CONTROLE NOTAS
    # ══════════════════════════════════════════════════════════
    df_dados = read_principal()
    df_hist = read_historico()

    pedido_in_dados = False
    pedido_in_hist = False
    status_anterior = None
    nf_existing = None

    if not df_dados.empty:
        df_dados = _ensure_nf_column(df_dados)
        match = df_dados[df_dados["PEDIDO"].astype(str).str.strip() == pedido_clean]
        if not match.empty:
            pedido_in_dados = True
            row_match = match.iloc[0]
            status_anterior = str(row_match.get("STATUS", "")).strip()
            nf_existing = row_match.get("NF")
            result["location"] = "Dados"

    if not pedido_in_dados and not df_hist.empty:
        match = df_hist[df_hist["PEDIDO"].astype(str).str.strip() == pedido_clean]
        if not match.empty:
            pedido_in_hist = True
            row_match = match.iloc[0]
            status_anterior = str(row_match.get("STATUS", "")).strip()
            nf_existing = row_match.get("NF")
            result["location"] = "Historico"

    # PEDIDO não encontrado → notificar
    if not pedido_in_dados and not pedido_in_hist:
        result["message"] = (
            f"⚠️ Pedido {pedido_clean} não encontrado na planilha CONTROLE NOTAS."
        )
        return result

    result["found_in_controle"] = True
    result["status_anterior"] = status_anterior

    # ══════════════════════════════════════════════════════════
    # PASSO 2: Verificar coluna NF da CONTROLE NOTAS
    # ══════════════════════════════════════════════════════════
    nf_existing_norm = _normalize_nf(nf_existing)
    result["nf_existing"] = nf_existing_norm or None

    if nf_existing_norm and nf_existing_norm != "":
        # Já tem NF anotada — confere com a do PDF?
        if nf_existing_norm == nf_clean:
            # Confere ✓ — segue para verificação no relatório
            result["nf_saved"] = True  # já estava anotada
        else:
            # CONFLITO — NF do PDF diverge da anotada
            result["nf_conflict"] = True
            result["message"] = (
                f"⚠️ Pedido {pedido_clean}: NF existente ({nf_existing_norm}) "
                f"diverge da NF do PDF ({nf_clean}). Verificar manualmente."
            )
            return result
    else:
        # NÃO tem NF → ANOTAR a NF do PDF na coluna NOTA FISCAL
        try:
            update_nf_batch({pedido_clean: nf_clean})
            result["nf_saved"] = True
            logger.info(f"NF {nf_clean} anotada para pedido {pedido_clean}")
        except Exception as e:
            logger.warning(f"Erro ao anotar NF: {e}")
            result["message"] = f"Erro ao anotar NF {nf_clean} no pedido {pedido_clean}: {e}"
            return result

    # ══════════════════════════════════════════════════════════
    # PASSO 3: Verificar NF no RELATÓRIO DE ENTREGAS
    # ══════════════════════════════════════════════════════════
    delivery_data = verify_delivery_in_report(nf_clean, remetente)
    result["found_in_report"] = delivery_data is not None
    result["delivery_data"] = delivery_data

    if delivery_data:
        # NF ENCONTRADA no relatório → foi ENTREGUE
        try:
            update_status_batch({pedido_clean: "ENTREGUE"})
            result["status_updated"] = True
        except Exception as e:
            result["success"] = True
            result["message"] = (
                f"📝 NF {nf_clean} anotada e encontrada no relatório, "
                f"mas erro ao atualizar status: {e}"
            )
            return result

        # Salvar expedição retroativa
        try:
            _save_retroactive_expedition(pedido_clean, nf_clean, remetente, delivery_data)
        except Exception as e:
            logger.warning(f"Erro ao salvar expedição retroativa: {e}")

        result["success"] = True
        result["message"] = (
            f"✅ Pedido {pedido_clean} → ENTREGUE. "
            f"NF {nf_clean} encontrada no relatório "
            f"(veículo: {delivery_data.get('VEICULO', 'N/A')}, "
            f"data: {delivery_data['DATA'].strftime('%d/%m/%Y') if hasattr(delivery_data.get('DATA'), 'strftime') else 'N/A'})."
        )
    else:
        # NF NÃO encontrada no relatório → notificar, manter status
        result["success"] = True
        result["message"] = (
            f"📝 Pedido {pedido_clean}: NF {nf_clean} registrada na CONTROLE NOTAS. "
            f"NF não encontrada no relatório de entregas — "
            f"status mantido como '{status_anterior}'."
        )

    return result


def _save_retroactive_expedition(
    pedido: str, nf: str, remetente: str, delivery_data: Optional[dict]
):
    """Salva registro retroativo no expedicoes.xlsx usando append (sem apagar existentes)."""
    from modules.expedition_engine import append_expedition

    data_entrega = ""
    if delivery_data and delivery_data.get("DATA"):
        dt = delivery_data["DATA"]
        if hasattr(dt, "strftime"):
            data_entrega = dt.strftime("%d/%m/%Y")
        else:
            data_entrega = str(dt)

    expedition_item = {
        "DATA_EXPEDICAO": data_entrega or datetime.now().strftime("%d/%m/%Y"),
        "PEDIDO": pedido,
        "CLIENTE": delivery_data.get("CLIENTE", "") if delivery_data else "",
        "DESTINO": delivery_data.get("BAIRRO", "") if delivery_data else "",
        "VEICULO": delivery_data.get("VEICULO", "") if delivery_data else "",
        "MOTORISTA": "",
        "ORDEM_ENTREGA": "",
        "NF": nf,
        "PESO": delivery_data.get("PESO", 0) if delivery_data else 0,
        "VOLUMES": delivery_data.get("VOLUMES", 0) if delivery_data else 0,
        "CARREGADO": "SIM",
        "OBS": f"REVERSO - {remetente}",
        "CRIADO_EM": datetime.now().strftime("%d/%m/%Y %H:%M"),
    }

    try:
        append_expedition([expedition_item])
    except Exception as e:
        # Fallback: salva em CSV se expedicoes.xlsx falhar
        logger.warning(f"Falha no append_expedition, salvando em CSV: {e}")
        _save_expedition_csv_fallback(expedition_item)


def _save_expedition_csv_fallback(row: dict):
    """Fallback: salva registro em CSV separado."""
    import csv

    csv_path = os.path.join(
        os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
        "confirmacoes_reversas.csv",
    )
    file_exists = os.path.exists(csv_path)
    with open(csv_path, "a", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=row.keys())
        if not file_exists:
            writer.writeheader()
        writer.writerow(row)


# ══════════════════════════════════════════════════════════════
# BATCH — PROCESSAR MÚLTIPLAS DANFEs
# ══════════════════════════════════════════════════════════════

def batch_process_danfes(pdf_files: list) -> list[dict]:
    """
    Processa múltiplas DANFEs em lote:
    1. Para cada DANFE: extrai NF + Pedido_Associado + Remetente
    2. Retorna lista de extrações para revisão do usuário

    Args:
        pdf_files: Lista de file-like objects (UploadedFiles do Streamlit)

    Returns:
        Lista de dicts com dados extraídos de cada DANFE
    """
    from modules.danfe_parser import extrair_danfe

    results = []

    for pdf in pdf_files:
        nome = getattr(pdf, "name", "desconhecido")
        try:
            data = extrair_danfe(pdf)
            if data:
                results.append({
                    "arquivo": nome,
                    "nf": _normalize_nf(data.get("Nota_Fiscal", "")),
                    "pedido": str(data.get("Pedido", "")).strip(),
                    "remetente": data.get("Remetente", ""),
                    "cliente": data.get("Cliente", ""),
                    "data_emissao": data.get("Data", ""),
                    "peso": data.get("Peso", ""),
                    "volumes": data.get("Volumes", ""),
                    "valor_nota": data.get("Valor_Nota", ""),
                    "status": "extraido",
                    "erro": "",
                })
            else:
                results.append({
                    "arquivo": nome,
                    "nf": "",
                    "pedido": "",
                    "remetente": "",
                    "cliente": "",
                    "data_emissao": "",
                    "peso": "",
                    "volumes": "",
                    "valor_nota": "",
                    "status": "erro",
                    "erro": "Extração vazia",
                })
        except Exception as e:
            results.append({
                "arquivo": nome,
                "nf": "",
                "pedido": "",
                "remetente": "",
                "cliente": "",
                "data_emissao": "",
                "peso": "",
                "volumes": "",
                "valor_nota": "",
                "status": "erro",
                "erro": str(e),
            })

    return results


# ══════════════════════════════════════════════════════════════
# CONSULTAS
# ══════════════════════════════════════════════════════════════

def _ensure_nf_column(df: pd.DataFrame) -> pd.DataFrame:
    """Garante que a coluna NF existe no DataFrame (cache antigo pode não ter)."""
    if "NF" not in df.columns:
        # Cache antigo — limpa cache e recarrega
        import streamlit as st
        st.cache_data.clear()
        from modules.excel_handler import read_principal
        df = read_principal()
        # Se ainda não tem NF, adiciona coluna vazia
        if "NF" not in df.columns:
            df["NF"] = None
    return df


def get_pending_orders_without_nf() -> pd.DataFrame:
    """
    Retorna pedidos que estão em status ativo mas ainda sem NF.

    Filtra por status: CONCLUIDO, AGUARDANDO NF, EM ROTA, SEPARACAO, PARCIAL
    """
    from modules.excel_handler import read_principal

    df = read_principal()
    if df.empty:
        return df

    df = _ensure_nf_column(df)

    # Status que indicam pedidos pendentes de confirmação
    active_statuses = [
        "SEPARACAO", "PARCIAL", "CONCLUIDO",
        "AGUARDANDO NF", "EM ROTA", "ENVIAR DATA",
    ]

    mask_status = df["STATUS"].isin(active_statuses)
    mask_no_nf = (
        df["NF"].isna()
        | (df["NF"].astype(str).str.strip() == "")
        | (df["NF"].astype(str).str.lower().isin(["nan", "none", "0"]))
    )

    return df[mask_status & mask_no_nf].copy()


def get_orders_with_nf_pending_delivery() -> pd.DataFrame:
    """
    Retorna pedidos que já têm NF mas ainda não estão ENTREGUE.
    Útil para verificação em lote no relatório.
    """
    from modules.excel_handler import read_principal

    df = read_principal()
    if df.empty:
        return df

    df = _ensure_nf_column(df)

    mask_has_nf = (
        df["NF"].notna()
        & (df["NF"].astype(str).str.strip() != "")
        & (~df["NF"].astype(str).str.lower().isin(["nan", "none", "0"]))
    )
    mask_not_delivered = df["STATUS"] != "ENTREGUE"

    return df[mask_has_nf & mask_not_delivered].copy()


# ══════════════════════════════════════════════════════════════
# BATCH OTIMIZADO — CONFIRMAR MÚLTIPLOS PEDIDOS EM UMA PASSADA
# ══════════════════════════════════════════════════════════════

def confirm_delivery_batch(items: list[dict]) -> list[dict]:
    """
    Versão otimizada de confirm_delivery para processar múltiplos pedidos.

    Em vez de abrir/ler/salvar o Excel N vezes, faz:
    1. Lê CONTROLE NOTAS (Dados) UMA VEZ
    2. Lê RELATÓRIO DE ENTREGAS UMA VEZ
    3. Processa todos os pedidos em memória
    4. Grava todas as NFs em UMA operação batch
    5. Grava todos os status em UMA operação batch
    6. Append de todas as expedições em UMA operação batch

    Performance: O(1) de I/O independente do tamanho do lote.

    Args:
        items: Lista de dicts com keys: "pedido", "nf", "remetente"

    Returns:
        Lista de dicts com resultado de cada pedido.
    """
    from modules.excel_handler import (
        read_principal,
        read_historico,
        update_nf_batch,
        update_status_batch,
    )
    from modules.expedition_engine import append_expedition

    if not items:
        return []

    # ══════════════════════════════════════════════════════════
    # FASE 1: LEITURA — tudo de uma vez
    # ══════════════════════════════════════════════════════════
    df_dados = read_principal()
    df_dados = _ensure_nf_column(df_dados)

    # Cria index para busca rápida O(1) por pedido
    pedido_index = {}
    if not df_dados.empty:
        for idx, row in df_dados.iterrows():
            pedido_str = str(row.get("PEDIDO", "")).strip().lstrip('0')
            if pedido_str:
                pedido_index[pedido_str] = {
                    "STATUS": str(row.get("STATUS", "")).strip(),
                    "NF": row.get("NF"),
                    "location": "Dados",
                }

    # Também lê aba Historico (pedidos arquivados)
    try:
        df_hist = read_historico()
        if not df_hist.empty:
            df_hist = _ensure_nf_column(df_hist)
            for idx, row in df_hist.iterrows():
                pedido_str = str(row.get("PEDIDO", "")).strip().lstrip('0')
                # Só adiciona se não já existir em Dados
                if pedido_str and pedido_str not in pedido_index:
                    pedido_index[pedido_str] = {
                        "STATUS": str(row.get("STATUS", "")).strip(),
                        "NF": row.get("NF"),
                        "location": "Historico",
                    }
    except Exception as e:
        logger.warning(f"Não foi possível ler aba Historico: {e}")

    # Lê relatório de entregas UMA VEZ e prepara para buscas rápidas
    df_report = pd.DataFrame()
    try:
        from modules.delivery_reader import read_deliveries_report
        df_report = read_deliveries_report()
        if not df_report.empty:
            df_report["_NF_NORM"] = df_report["NOTA_FISCAL"].apply(
                lambda x: str(int(x)) if pd.notna(x) and x != 0 else ""
            )
            df_report["_REM_NORM"] = df_report["REMETENTE"].apply(_normalize_remetente)
    except Exception as e:
        logger.warning(f"Não foi possível ler relatório de entregas: {e}")


    # ══════════════════════════════════════════════════════════
    # FASE 2: PROCESSAMENTO — tudo em memória
    # ══════════════════════════════════════════════════════════
    results = []
    nf_updates = {}       # {pedido: nf} — para gravar em batch
    status_updates = {}   # {pedido: "ENTREGUE"} — para gravar em batch
    expedition_items = [] # lista de dicts para append em batch

    for item in items:
        pedido_clean = str(item.get("pedido", "")).strip().lstrip('0')
        nf_clean = _normalize_nf(item.get("nf", ""))
        remetente = str(item.get("remetente", "")).strip()

        result = {
            "success": False,
            "pedido": pedido_clean,
            "nf": nf_clean,
            "found_in_controle": False,
            "nf_existing": None,
            "nf_conflict": False,
            "nf_saved": False,
            "found_in_report": False,
            "delivery_data": None,
            "status_updated": False,
            "location": None,
            "status_anterior": None,
            "message": "",
        }

        # ── Busca pedido (O(1) no index) ──
        info = pedido_index.get(pedido_clean)
        if not info:
            result["message"] = (
                f"⚠️ Pedido {pedido_clean} não encontrado na planilha CONTROLE NOTAS."
            )
            results.append(result)
            continue

        result["found_in_controle"] = True
        result["location"] = info["location"]
        result["status_anterior"] = info["STATUS"]

        # ── Verifica NF existente ──
        nf_existing_norm = _normalize_nf(info["NF"])
        result["nf_existing"] = nf_existing_norm or None

        if nf_existing_norm and nf_existing_norm != "":
            if nf_existing_norm == nf_clean:
                result["nf_saved"] = True
            else:
                result["nf_conflict"] = True
                result["success"] = True
                result["message"] = (
                    f"⚠️ Pedido {pedido_clean}: NF existente ({nf_existing_norm}) "
                    f"diverge da NF do PDF ({nf_clean}). Verificar manualmente."
                )
                results.append(result)
                continue
        else:
            nf_updates[pedido_clean] = nf_clean
            result["nf_saved"] = True

        # ── Busca NF no relatório (usando DataFrame pré-carregado) ──
        delivery_data = _search_report_fast(df_report, nf_clean, remetente)
        result["found_in_report"] = delivery_data is not None
        result["delivery_data"] = delivery_data

        if delivery_data:
            status_updates[pedido_clean] = "ENTREGUE"
            result["status_updated"] = True
            result["success"] = True

            # Prepara item de expedição
            data_entrega = ""
            dt = delivery_data.get("DATA")
            if dt and hasattr(dt, "strftime"):
                data_entrega = dt.strftime("%d/%m/%Y")

            expedition_items.append({
                "DATA_EXPEDICAO": data_entrega or datetime.now().strftime("%d/%m/%Y"),
                "PEDIDO": pedido_clean,
                "CLIENTE": delivery_data.get("CLIENTE", ""),
                "DESTINO": delivery_data.get("BAIRRO", ""),
                "VEICULO": delivery_data.get("VEICULO", ""),
                "MOTORISTA": "",
                "ORDEM_ENTREGA": "",
                "NF": nf_clean,
                "PESO": delivery_data.get("PESO", 0),
                "VOLUMES": delivery_data.get("VOLUMES", 0),
                "CARREGADO": "SIM",
                "OBS": f"REVERSO - {remetente}",
                "CRIADO_EM": datetime.now().strftime("%d/%m/%Y %H:%M"),
            })

            result["message"] = (
                f"✅ Pedido {pedido_clean} → ENTREGUE. "
                f"NF {nf_clean} encontrada no relatório "
                f"(veículo: {delivery_data.get('VEICULO', 'N/A')}, "
                f"data: {data_entrega or 'N/A'})."
            )
        else:
            result["success"] = True
            result["message"] = (
                f"📝 Pedido {pedido_clean}: NF {nf_clean} registrada na CONTROLE NOTAS. "
                f"NF não encontrada no relatório de entregas — "
                f"status mantido como '{info['STATUS']}'."
            )

        results.append(result)

    # ══════════════════════════════════════════════════════════
    # FASE 3: GRAVAÇÃO — tudo de uma vez
    # ══════════════════════════════════════════════════════════
    errors = []

    # 3a. Gravar NFs em batch (1 open + 1 save)
    if nf_updates:
        try:
            update_nf_batch(nf_updates)
            logger.info(f"NF batch: {len(nf_updates)} atualizadas")
        except Exception as e:
            errors.append(f"Erro ao gravar NFs: {e}")
            logger.error(f"Erro update_nf_batch: {e}")

    # 3b. Gravar status em batch (1 open + 1 save)
    if status_updates:
        try:
            update_status_batch(status_updates)
            logger.info(f"Status batch: {len(status_updates)} atualizados")
        except Exception as e:
            errors.append(f"Erro ao atualizar status: {e}")
            logger.error(f"Erro update_status_batch: {e}")
            # Marca como não atualizado
            for r in results:
                if r["pedido"] in status_updates:
                    r["status_updated"] = False

    # 3c. Append expedições em batch (1 open + 1 save)
    if expedition_items:
        try:
            append_expedition(expedition_items)
            logger.info(f"Expedições: {len(expedition_items)} adicionadas")
        except Exception as e:
            logger.warning(f"Erro ao salvar expedições: {e}")

    # Adiciona erros globais na mensagem do primeiro resultado
    if errors and results:
        results[0]["message"] += " | ERROS: " + "; ".join(errors)

    return results


def _search_report_fast(
    df_report: pd.DataFrame, nf: str, remetente: str
) -> Optional[dict]:
    """
    Busca rápida em DataFrame pré-carregado do relatório (sem re-leitura).
    Usa as colunas _NF_NORM e _REM_NORM já preparadas.
    """
    if df_report.empty or not nf:
        return None

    rem_norm = _normalize_remetente(remetente)

    # Busca: REMETENTE + NF (chave composta — NFs são únicas por remetente)
    matches = pd.DataFrame()
    if rem_norm:
        mask = (df_report["_NF_NORM"] == nf) & (df_report["_REM_NORM"] == rem_norm)
        matches = df_report[mask]

    if matches.empty:
        return None

    row = matches.iloc[0]
    return {
        "DATA": row.get("DATA"),
        "VEICULO": row.get("VEICULO", ""),
        "OPERACAO": row.get("OPERACAO", ""),
        "REMETENTE": row.get("REMETENTE", ""),
        "CLIENTE": row.get("CLIENTE", ""),
        "BAIRRO": row.get("BAIRRO", ""),
        "UF": row.get("UF", ""),
        "NOTA_FISCAL": nf,
        "PESO": row.get("PESO", 0),
        "VOLUMES": row.get("VOLUMES", 0),
        "VALOR_NOTA": row.get("VALOR_NOTA", 0),
        "FRETE": row.get("FRETE", 0),
    }
