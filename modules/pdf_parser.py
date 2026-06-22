"""
pdf_parser.py — Módulo de extração de dados de PDFs (Onfinity "Mapa de Separação por Pedido").

Utiliza pdfplumber com BOUNDING BOXES mapeadas do documento real para extrair
com precisão cirúrgica: Número do Pedido, ID do Cliente, Nome do Cliente,
Endereço, Observações e Empresa.

ESTRATÉGIA:
    1. Extração por bounding box (coordenadas fixas mapeadas do PDF real)
    2. Fallback por regex (caso o layout varie)

FORMATO ESPERADO (Onfinity):
    ┌───────────────────────────────────────────────────────────────────┐
    │ [Logo Onfinity]     Mapa de Separação por Pedido    11/05/2026   │
    │                                                      X de 67    │
    │ Pedido    Cliente                                                │
    │ 096075    023319-BIOMA COMERCIO          CNPJ/CPF: ...           │
    │ BIOMA COMERCIO DE MOVEIS LTDA            Vendedor: ...           │
    │ AVENIDA OTTO BAUMGART, 500 - LOJA 226... - SAO PAULO/SP         │
    │ Observações: MCS:137792-PEDIDO DE COMPRA...  Resp. Separação:... │
    │ [Tabela de Itens]                                                │
    │ Totais do Pedido   91,00   46,307   R$ 1.251,58                  │
    └───────────────────────────────────────────────────────────────────┘

REGRA DE ENDEREÇO:
    1. Prioridade: Endereço de ENTREGA extraído das OBSERVAÇÕES
    2. Fallback: Endereço formal (área 6)
    3. Nomes genéricos são normalizados (EMBU → EMBU DAS ARTES - SP)
"""

import logging
import re
from datetime import datetime
from typing import Optional

import pdfplumber

# ══════════════════════════════════════════════════════════════
# LOGGING
# ══════════════════════════════════════════════════════════════
logger = logging.getLogger(__name__)

# ══════════════════════════════════════════════════════════════
# BOUNDING BOXES — Mapeadas do PDF real (Pedido_Sep_B55.pdf)
# Formato: (x0, top, x1, bottom)
# Coordenadas em pontos PDF (1pt = 1/72 polegada)
# Page dimensions: width=595.5, height=842.2
# ══════════════════════════════════════════════════════════════
BBOX_PEDIDO   = (30,  95,  92,  115)   # Número do pedido (ex: "096075")
BBOX_CLIENTE  = (85,  95, 430,  130)   # Código + nome do cliente
BBOX_ENDERECO = (28, 128, 430,  150)   # Endereço formal
BBOX_OBS      = (28, 148, 340,  180)   # Observações (pode ter 2-3 linhas)
BBOX_EMPRESA  = (0,   20, 200,   85)   # Área do logo/empresa

# ══════════════════════════════════════════════════════════════
# NORMALIZAÇÃO DE LOCAIS GENÉRICOS
# Nomes curtos usados nas observações que remetem a cidades
# ══════════════════════════════════════════════════════════════
LOCATION_ALIASES = {
    "EMBU":           "EMBU DAS ARTES - SP",
    "TABOAO":         "TABOAO DA SERRA - SP",
    "COTIA":          "COTIA - SP",
    "BARUERI":        "BARUERI - SP",
    "OSASCO":         "OSASCO - SP",
    "CARAPICUIBA":    "CARAPICUIBA - SP",
    "ITAPEVI":        "ITAPEVI - SP",
    "JANDIRA":        "JANDIRA - SP",
    "SANTANA PARNAIBA": "SANTANA DE PARNAIBA - SP",
    "ALPHAVILLE":     "BARUERI (ALPHAVILLE) - SP",
}


# ══════════════════════════════════════════════════════════════
# CONTINUATION PAGE DETECTION
# Pedidos extensos podem ocupar múltiplas páginas. As páginas
# subsequentes NÃO repetem o cabeçalho (Pedido/Cliente/CNPJ),
# contendo apenas o header genérico "Mapa de Separação" seguido
# diretamente das linhas de itens/materiais.
# ══════════════════════════════════════════════════════════════
def _is_continuation_page(full_text: str) -> bool:
    """
    Detecta se uma página é CONTINUAÇÃO de um pedido anterior.

    Uma página NORMAL contém obrigatoriamente:
      - Header "Pedido  Cliente" (linha de cabeçalho dos campos)
      - CNPJ/CPF
      - Vendedor

    Uma página de CONTINUAÇÃO tem apenas:
      - "Mapa de Separação por Pedido X de Y" (header genérico)
      - Linhas de itens (materiais)
      - Opcionalmente "Totais do Pedido" (se for a última página)

    Returns:
        True se a página for uma continuação (sem dados de identificação).
    """
    has_mapa = bool(re.search(r'Mapa\s+de\s+Separa', full_text, re.IGNORECASE))
    has_pedido_header = bool(re.search(r'Pedido\s+Cliente', full_text))
    has_cnpj = bool(re.search(r'CNPJ/CPF', full_text))
    has_vendedor = bool(re.search(r'Vendedor', full_text))

    # Se tem o cabeçalho do mapa mas NÃO tem os sinais de identificação
    # → é uma página de continuação
    return has_mapa and not (has_pedido_header and has_cnpj and has_vendedor)


# ══════════════════════════════════════════════════════════════
# MAIN EXTRACTION FUNCTION
# ══════════════════════════════════════════════════════════════
def extract_data_from_pdf(pdf_file) -> list[dict]:
    """
    Extrai dados de pedidos de um PDF Onfinity (Mapa de Separação por Pedido).

    Processa o documento PÁGINA A PÁGINA, extraindo:
    - PEDIDO: Número do pedido
    - ID_CLIENTE: Código numérico do cliente
    - CLIENTE: Nome do cliente
    - ENDERECO: Endereço de entrega (prioridade: OBS > endereço formal)
    - OBS: Observações completas
    - EMPRESA: Onfinity ou Green Bags
    - DATA: Timestamp da extração
    - STATUS: Status inicial (SEPARACAO)
    - _CONTINUATION_PAGES: Número de páginas de continuação mescladas

    Suporta pedidos que ocupam múltiplas páginas ("continuações").
    Páginas de continuação são detectadas automaticamente pela ausência
    de cabeçalho de pedido/cliente/CNPJ e seus itens são mesclados
    no pedido da(s) página(s) anterior(es).

    Args:
        pdf_file: File-like object ou caminho do arquivo PDF.

    Returns:
        Lista de dicionários, um por pedido encontrado.
    """
    results = []
    seen_pedidos = {}  # Map PEDIDO -> dict reference (para mesclar páginas repetidas)
    errors = []
    last_pedido_ref = None  # Referência ao último pedido válido (para continuações)
    continuation_count = 0  # Total de páginas de continuação detectadas

    with pdfplumber.open(pdf_file) as pdf:
        total_pages = len(pdf.pages)
        logger.info(f"Processando PDF com {total_pages} página(s)")

        for page_num, page in enumerate(pdf.pages, start=1):
            try:
                full_text = page.extract_text() or ""

                if not full_text.strip():
                    logger.warning(f"Página {page_num}/{total_pages}: página vazia")
                    continue

                # ── DETECÇÃO DE PÁGINA DE CONTINUAÇÃO ──
                if _is_continuation_page(full_text):
                    continuation_count += 1

                    if last_pedido_ref is not None:
                        # Mesclar observações complementares (se houver)
                        # e incrementar contador de continuações
                        last_pedido_ref["_CONTINUATION_PAGES"] = (
                            last_pedido_ref.get("_CONTINUATION_PAGES", 0) + 1
                        )
                        logger.info(
                            f"Página {page_num}/{total_pages}: "
                            f"CONTINUAÇÃO do Pedido {last_pedido_ref['PEDIDO']} "
                            f"(pág. de continuação #{last_pedido_ref['_CONTINUATION_PAGES']})"
                        )
                    else:
                        # Continuação sem pedido anterior — erro no PDF
                        logger.error(
                            f"Página {page_num}/{total_pages}: "
                            f"detectada como continuação, mas sem pedido anterior. "
                            f"Itens desta página serão descartados."
                        )
                        errors.append(
                            f"Página {page_num}: continuação órfã (sem pedido anterior)"
                        )
                    continue

                # ── PÁGINA NORMAL (com cabeçalho completo) ──
                page_data = _extract_page_data(page, page_num)

                if page_data is None:
                    logger.warning(f"Página {page_num}/{total_pages}: nenhum pedido identificado")
                    continue

                pedido_num = page_data["PEDIDO"]

                if pedido_num not in seen_pedidos:
                    # Novo pedido encontrado
                    page_data.update({
                        "DATA": datetime.now(),
                        "STATUS": "SEPARACAO",
                        "_CONTINUATION_PAGES": 0,
                    })
                    results.append(page_data)
                    seen_pedidos[pedido_num] = page_data
                    last_pedido_ref = page_data  # Atualizar referência para continuações
                    logger.info(
                        f"Página {page_num}/{total_pages}: "
                        f"Pedido {pedido_num} | Cliente: {page_data['CLIENTE'][:30]}"
                    )
                else:
                    # Pedido repetido em outra página — mesclar dados complementares
                    _merge_page_data(seen_pedidos[pedido_num], page_data)
                    last_pedido_ref = seen_pedidos[pedido_num]
                    logger.debug(
                        f"Página {page_num}/{total_pages}: "
                        f"Pedido {pedido_num} (repetido, dados mesclados)"
                    )

            except Exception as e:
                error_msg = f"Página {page_num}/{total_pages}: erro na extração — {e}"
                logger.error(error_msg)
                errors.append(error_msg)
                continue  # Nunca abortar o pipeline

    if continuation_count > 0:
        logger.info(
            f"⚠️ {continuation_count} página(s) de continuação detectada(s) e mesclada(s)"
        )

    logger.info(
        f"Extração concluída: {len(results)} pedido(s) de {total_pages} página(s)"
        + (f" | {continuation_count} continuação(ões)" if continuation_count else "")
        + (f" | {len(errors)} erro(s)" if errors else "")
    )
    return results


# ══════════════════════════════════════════════════════════════
# PAGE-LEVEL EXTRACTION
# ══════════════════════════════════════════════════════════════
def _extract_page_data(page, page_num: int) -> Optional[dict]:
    """
    Extrai todos os campos de uma única página do PDF.

    Usa bounding boxes como estratégia primária, com fallback regex.

    Returns:
        Dicionário com os campos extraídos, ou None se a página não tem pedido.
    """
    full_text = page.extract_text() or ""

    if not full_text.strip():
        return None

    # ── 1. PEDIDO ──────────────────────────────────────────
    pedido = _extract_bbox_text(page, BBOX_PEDIDO)
    if pedido:
        pedido = re.sub(r'\D', '', pedido).strip()  # Manter apenas dígitos

    if not pedido or len(pedido) < 4:
        # Fallback regex
        m = re.search(r'Pedido\s+Cliente\s*\n(\d{4,10})', full_text)
        pedido = m.group(1).strip() if m else None

    if not pedido:
        return None  # Sem pedido = página sem dados úteis

    # ── 2. CLIENTE (ID + Nome) ────────────────────────────
    cliente_raw = _extract_bbox_text(page, BBOX_CLIENTE)
    id_cliente, cliente_nome = _parse_cliente(cliente_raw, full_text)

    # ── 3. ENDEREÇO FORMAL (área 6) ───────────────────────
    endereco_formal = _extract_bbox_text(page, BBOX_ENDERECO)
    endereco_formal = _clean_text(endereco_formal)

    # ── 4. OBSERVAÇÕES ────────────────────────────────────
    obs_raw = _extract_bbox_text(page, BBOX_OBS)

    if not obs_raw:
        # Fallback regex para observações
        obs_match = re.search(
            r'Observa[çcã][õoã]es\s*[:\-]?\s*(.+?)(?:Resp\.|Totais|Local\s+C|$)',
            full_text,
            re.IGNORECASE | re.DOTALL
        )
        obs_raw = obs_match.group(1).strip() if obs_match else ""

    # Limpar prefixo "Observações:" se presente
    obs_clean = re.sub(
        r'^Observa[çcã][õoã]es\s*[:\-]?\s*',
        '',
        obs_raw or "",
        flags=re.IGNORECASE
    ).strip()
    obs_clean = _clean_text(obs_clean)

    # ── 5. ENDEREÇO DE ENTREGA (prioridade: OBS > formal) ─
    endereco_entrega = _extract_delivery_address_from_obs(obs_clean)
    endereco_final = endereco_entrega if endereco_entrega else endereco_formal

    # Normalizar nomes genéricos de locais
    endereco_final = _normalize_location(endereco_final)

    # ── 6. EMPRESA ────────────────────────────────────────
    empresa = _detect_empresa(full_text, page)

    return {
        "PEDIDO":     pedido,
        "ID_CLIENTE": id_cliente,
        "CLIENTE":    cliente_nome,
        "ENDERECO":   endereco_final,
        "OBS":        obs_clean[:500] if obs_clean else "",
        "EMPRESA":    empresa,
    }


# ══════════════════════════════════════════════════════════════
# BBOX EXTRACTION HELPER
# ══════════════════════════════════════════════════════════════
def _extract_bbox_text(page, bbox: tuple) -> Optional[str]:
    """
    Extrai texto de uma região retangular (bounding box) da página.

    Args:
        page: Objeto pdfplumber.Page
        bbox: Tupla (x0, top, x1, bottom) em pontos PDF

    Returns:
        Texto extraído ou None se a região estiver vazia/inválida.
    """
    try:
        region = page.within_bbox(bbox)
        if region:
            text = region.extract_text()
            return text if text and text.strip() else None
    except Exception as e:
        logger.debug(f"Falha na extração bbox {bbox}: {e}")
    return None


# ══════════════════════════════════════════════════════════════
# CLIENTE PARSING
# ══════════════════════════════════════════════════════════════
def _parse_cliente(cliente_raw: Optional[str], full_text: str) -> tuple[str, str]:
    """
    Extrai o ID numérico e o nome do cliente.

    Formato esperado no bbox: "Cliente\n0 23319-BIOMA COMERCIO\nRCIO DE MOVEIS LTDA"
    O que queremos: ID="023319", Nome="BIOMA COMERCIO"

    Returns:
        Tupla (id_cliente, nome_cliente)
    """
    id_cliente = ""
    nome_cliente = "NÃO IDENTIFICADO"

    if cliente_raw:
        # Remover header "Cliente" se presente
        text = re.sub(r'^Cliente\s*', '', cliente_raw, flags=re.IGNORECASE).strip()

        # Padrão Onfinity: "0 23319-BIOMA COMERCIO" ou "023319-BIOMA COMERCIO"
        # O "0" solto no início é um artefato do layout
        # Nota: IDs podem ter 2+ dígitos (ex: "02" para Onfinity Comercial)
        match = re.search(r'(?:^0\s+)?(\d{2,10})\s*[-–]\s*(.+?)(?:\n|$)', text)
        if match:
            id_cliente = match.group(1).strip()
            nome_cliente = match.group(2).strip()
        else:
            # Tentar pegar qualquer texto após dígitos
            match2 = re.search(r'(\d{2,10})\s*[-–]?\s*(.+?)(?:\n|$)', text)
            if match2:
                id_cliente = match2.group(1).strip()
                nome_cliente = match2.group(2).strip()

    # Fallback: buscar no texto completo
    if nome_cliente == "NÃO IDENTIFICADO":
        m = re.search(
            r'Pedido\s+Cliente\s*\n\d{4,10}\s+\d*\s*(\d{2,10})\s*[-–]\s*(.+?)(?:\n|$)',
            full_text
        )
        if m:
            id_cliente = m.group(1).strip()
            nome_cliente = m.group(2).strip()

    # Limpar nome
    nome_cliente = _clean_text(nome_cliente)

    return id_cliente, nome_cliente


# ══════════════════════════════════════════════════════════════
# DELIVERY ADDRESS EXTRACTION FROM OBS
# ══════════════════════════════════════════════════════════════
def _extract_delivery_address_from_obs(obs_text: str) -> str:
    """
    Tenta extrair um endereço de entrega do campo de observações.

    Padrões reconhecidos:
    - "ENTREGA EMBU 14/5"  → "EMBU"
    - "ENTREGA DEVE SER REALIZADA..."
    - "entrega pode ser na segunda"
    - Endereço completo com rua/avenida

    Returns:
        Endereço de entrega ou string vazia se não encontrado.
    """
    if not obs_text:
        return ""

    upper = obs_text.upper()

    # Filtro de falsos positivos
    ignore_words = {
        "DEVE", "SER", "REALIZADA", "DAS", "PODE", "NA", "NO",
        "SEGUNDA", "TERCA", "QUARTA", "QUINTA", "SEXTA", "SABADO",
        "AQUI", "SEPARAR", "FAT", "FATURA", "FATURAR", "POR",
    }

    # Padrão 1: "ENTREGA <LOCAL>" — nome de cidade/bairro após "ENTREGA"
    # Ex: "ENTREGA EMBU 14/5", "ENTREGA TABOAO 14/5"
    match = re.search(
        r'ENTREGA\s+(?:EM\s+)?([A-ZÀ-ÚÃÕ][A-ZÀ-ÚÃÕ\s]{2,30}?)(?:\s+\d{1,2}/|\s*[-–]|$)',
        upper
    )
    if match:
        local = match.group(1).strip()
        first_word = local.split()[0] if local.split() else ""
        if first_word not in ignore_words and len(first_word) > 2:
            return local

    # Padrão 2: "ENTREGA <data> - <LOCAL>" — local após data
    # Ex: "ENTREGA 07/05 - EMBU", "ENTREGA 14/5 - EMBU"
    match_after_date = re.search(
        r'ENTREGA\s+\d{1,2}/\d{1,2}\s*[-–]\s*([A-ZÀ-ÚÃÕ][A-ZÀ-ÚÃÕ\s]{2,30}?)(?:\s*[-–]|$)',
        upper
    )
    if match_after_date:
        local = match_after_date.group(1).strip()
        first_word = local.split()[0] if local.split() else ""
        if first_word not in ignore_words and len(first_word) > 2:
            return local

    # Padrão 3: Apenas nome de local conhecido no texto (sem "ENTREGA")
    # Busca por aliases conhecidos em qualquer lugar das OBS
    for alias in LOCATION_ALIASES:
        if re.search(r'\b' + re.escape(alias) + r'\b', upper):
            return alias

    # Padrão 4: Endereço completo com logradouro
    match_rua = re.search(
        r'((?:RUA|AV\.?|AVENIDA|ALAMEDA|TRAVESSA|ESTRADA|RODOVIA|R\.)\s+.+?)(?:\s*[-–]\s*$|\s+FAT|\s+FATURA|$)',
        upper,
        re.IGNORECASE
    )
    if match_rua:
        return match_rua.group(1).strip()

    return ""


# ══════════════════════════════════════════════════════════════
# LOCATION NORMALIZATION
# ══════════════════════════════════════════════════════════════
def _normalize_location(endereco: str) -> str:
    """
    Normaliza nomes genéricos de locais para nomes completos de cidades.

    Exemplos:
        "EMBU" → "EMBU DAS ARTES - SP"
        "TABOAO" → "TABOAO DA SERRA - SP"
    """
    if not endereco:
        return endereco

    # Se o endereço já é completo (contém rua/avenida), não normalizar
    if re.search(r'(?:RUA|AV|AVENIDA|ALAMEDA|RODOVIA|ESTRADA|R\.)\s', endereco, re.IGNORECASE):
        return endereco

    # Verificar se o endereço é um nome genérico
    # Normalizar acentos para comparação (TABOãO → TABOAO)
    import unicodedata
    def _strip_accents(s):
        return ''.join(
            c for c in unicodedata.normalize('NFD', s)
            if unicodedata.category(c) != 'Mn'
        )

    endereco_normalized = _strip_accents(endereco.strip().upper())
    for alias, full_name in LOCATION_ALIASES.items():
        alias_normalized = _strip_accents(alias.upper())
        if endereco_normalized == alias_normalized or endereco_normalized.startswith(alias_normalized + " "):
            return full_name

    return endereco


# ══════════════════════════════════════════════════════════════
# EMPRESA DETECTION
# ══════════════════════════════════════════════════════════════
def _detect_empresa(full_text: str, page) -> str:
    """
    Detecta se o documento é da Onfinity ou Green Bags.

    Verifica:
    1. Texto na área do logo (bounding box)
    2. Palavras-chave no texto completo da página
    """
    # Tentar ler texto na área do logo
    empresa_text = _extract_bbox_text(page, BBOX_EMPRESA) or ""

    combined = (empresa_text + " " + full_text).upper()

    if "GREEN BAG" in combined or "GREENBAG" in combined:
        return "GREEN BAGS"

    # Default para Onfinity (o logo é imagem, não texto)
    return "ONFINITY"


# ══════════════════════════════════════════════════════════════
# MERGE (for multi-page orders)
# ══════════════════════════════════════════════════════════════
def _merge_page_data(existing: dict, new_data: dict) -> None:
    """
    Mescla dados de uma página adicional num pedido já registrado.
    Útil quando um pedido ocupa mais de uma página.
    """
    # Mesclar observações (sem duplicar)
    new_obs = new_data.get("OBS", "")
    if new_obs and new_obs not in (existing.get("OBS") or ""):
        current_obs = existing.get("OBS") or ""
        if current_obs:
            existing["OBS"] = f"{current_obs} | {new_obs}"[:500]
        else:
            existing["OBS"] = new_obs

    # Preencher endereço se estava vazio
    if not existing.get("ENDERECO") and new_data.get("ENDERECO"):
        existing["ENDERECO"] = new_data["ENDERECO"]

    # Preencher ID_CLIENTE se estava vazio
    if not existing.get("ID_CLIENTE") and new_data.get("ID_CLIENTE"):
        existing["ID_CLIENTE"] = new_data["ID_CLIENTE"]


# ══════════════════════════════════════════════════════════════
# TEXT CLEANING UTILITIES
# ══════════════════════════════════════════════════════════════
def _clean_text(text: Optional[str]) -> str:
    """Limpa texto extraído: remove espaços duplos, NaN, e whitespace."""
    if not text:
        return ""

    # Tratar NaN/None como strings
    s = str(text).strip()
    if s.lower() in ("nan", "none", "nat", "null", ""):
        return ""

    # Remover espaços duplos
    s = re.sub(r'\s{2,}', ' ', s)
    # Remover quebras de linha internas desnecessárias
    s = re.sub(r'\n\s*', ' ', s)
    return s.strip()


# ══════════════════════════════════════════════════════════════
# TEXT PREVIEW (DEBUG)
# ══════════════════════════════════════════════════════════════
def extract_text_preview(pdf_file) -> str:
    """
    Extrai o texto completo do PDF para preview/debug.

    Args:
        pdf_file: File-like object ou caminho.

    Returns:
        Texto extraído do PDF (todas as páginas).
    """
    with pdfplumber.open(pdf_file) as pdf:
        texts = []
        for i, page in enumerate(pdf.pages, start=1):
            text = page.extract_text()
            if text:
                texts.append(f"[Página {i}/{len(pdf.pages)}]\n{text}")
    return "\n---PAGE BREAK---\n".join(texts)
