"""
checklist_engine.py — Motor de cruzamento Estoque × Pedidos para Batch/Discrete Picking.

Responsável por:
- Ingestão de estoque (PDF CIGAM / Word .docx)
- Ingestão de pedidos (PDF Onfinity "Mapa de Separação por Pedido")
- Normalização de códigos e descrições
- Cruzamento e alocação virtual de estoque
- Geração de DataFrames para Batch e Discrete Picking
"""

import re
import io
import logging
import pandas as pd
import pdfplumber

logger = logging.getLogger(__name__)

# ══════════════════════════════════════════════════════════════
# LIMPEZA E NORMALIZAÇÃO
# ══════════════════════════════════════════════════════════════

def _limpar_quantidade(val) -> int:
    """Extrai inteiro de uma string de quantidade (formato BR: ponto=milhar, vírgula=decimal)."""
    if pd.isna(val) or not val:
        return 0
    s = str(val).replace('\n', '').strip()
    # Remove símbolos de moeda e espaços
    s = re.sub(r'[R$\s]', '', s)
    if not s:
        return 0
    # Formato BR: ponto é milhar, vírgula é decimal
    if ',' in s:
        s = s.replace('.', '').replace(',', '.')
    else:
        s = s.replace('.', '')
    try:
        return int(float(s))
    except (ValueError, TypeError):
        return 0


def _limpar_codigo(val) -> str:
    """Normaliza código SKU: remove espaços, newlines, converte para string."""
    if pd.isna(val) or not val:
        return ""
    return str(val).replace('\n', '').strip()


def _limpar_descricao(val) -> str:
    """Remove códigos de fornecedor e preserva código de referência (ex: 14891-COPO...)."""
    if pd.isna(val) or not val:
        return ""
    texto = str(val).replace('\n', ' ').strip()
    # Captura referência: DIGITOS-TEXTO (ex: "14891-COPO DESCARTAVEL", "C7138-DESODORIZADOR")
    m = re.search(r'(C?\d+\s*-\s*[A-Za-zÀ-ÚÃÕ#].+)', texto)
    if m:
        result = m.group(1).strip()
        # Normaliza espaços ao redor do hífen: "14891 - COPO" → "14891-COPO"
        result = re.sub(r'\s*-\s*', '-', result, count=1)
        return result.upper()
    return texto.strip().upper()


# ══════════════════════════════════════════════════════════════
# INGESTÃO DE ESTOQUE — PDF (CIGAM)
# ══════════════════════════════════════════════════════════════

def extrair_estoque_pdf(arquivo) -> pd.DataFrame:
    """
    Extrai posição de estoque do PDF CIGAM (Onfinity) via TEXTO.

    O PDF não possui bordas de tabela; pdfplumber fragmenta as colunas
    de forma inconsistente. A extração por TEXTO BRUTO é 100% confiável.

    Formato de cada linha de texto:
      CODIGO DESCRICAO UNI C.A. QUANTIDADE PR_MEDIO TOTAL_MEDIO
    Exemplo:
      6100050009 14896 - COPO DESCARTAVEL BRANCO 50ML PS KEROCOPO PCT 100UN PCT 001 1.054 0,0000 0,0000
    """
    dados = []
    arquivo.seek(0)

    # Regex para cada linha de item:
    # Grupo 1: Código (10 dígitos)
    # Grupo 2: Descrição (tudo entre código e unidade)
    # Grupo 3: Unidade (UN, PCT, CX, etc.)
    # Grupo 4: C.A. (3 dígitos)
    # Grupo 5: Quantidade (inteiro ou com ponto de milhar)
    RE_STOCK = re.compile(
        r'^(\d{10})\s+'                    # Código SKU (10 dígitos)
        r'(.+?)\s+'                         # Descrição (captura lazy)
        r'(UN|PCT|CX|RL|KG|MT|PR|DZ|CT|MH|LT|GL|SC|BD|FD|RS|TB|JG|PC|PAR|BL|GR|MIL)\s+'  # Unidade
        r'(\d{3})\s+'                       # C.A. (sempre 3 dígitos como "001")
        r'([\d.]+)'                         # Quantidade (pode ter ponto de milhar)
    )

    with pdfplumber.open(arquivo) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            for line in text.split('\n'):
                line = line.strip()
                if not line:
                    continue

                m = RE_STOCK.match(line)
                if m:
                    codigo = m.group(1).strip()
                    descricao_raw = m.group(2).strip()
                    qtd_str = m.group(5).strip()

                    # Quantidade: ponto é milhar no formato BR
                    qtd = _limpar_quantidade(qtd_str)

                    dados.append({
                        "Código": codigo,
                        "Descrição": _limpar_descricao(descricao_raw),
                        "Estoque_Atual": qtd,
                    })

    if not dados:
        return pd.DataFrame(columns=["Código", "Descrição", "Estoque_Atual"])

    df = pd.DataFrame(dados)
    # Agrupa duplicatas (mesmo código em páginas diferentes)
    df = df.groupby("Código", as_index=False).agg({
        "Descrição": "first",
        "Estoque_Atual": "sum"
    })
    return df


# ══════════════════════════════════════════════════════════════
# INGESTÃO DE ESTOQUE — Word (.docx)
# ══════════════════════════════════════════════════════════════

def extrair_estoque_docx(arquivo) -> pd.DataFrame:
    """Extrai estoque de arquivo Word (.docx) com tabelas."""
    import docx

    dados = []
    try:
        arquivo.seek(0)
        buf = io.BytesIO(arquivo.read())
        doc = docx.Document(buf)

        for table in doc.tables:
            header_found = False
            col_map = {}
            for i, row in enumerate(table.rows):
                cells = [c.text.replace('\n', ' ').strip() for c in row.cells]
                if not header_found:
                    # Detecta header
                    upper_cells = [c.upper() for c in cells]
                    for j, c in enumerate(upper_cells):
                        if 'CÓDIGO' in c or 'CODIGO' in c:
                            col_map['codigo'] = j
                        if 'DESCRI' in c:
                            col_map['descricao'] = j
                        if 'QUANT' in c:
                            col_map['quantidade'] = j
                    if 'codigo' in col_map:
                        header_found = True
                    continue

                idx_cod = col_map.get('codigo', 0)
                idx_desc = col_map.get('descricao', 1)
                idx_qtd = col_map.get('quantidade', 4)

                if idx_cod < len(cells):
                    codigo = _limpar_codigo(cells[idx_cod])
                else:
                    continue
                if not codigo or not re.match(r'^\d', codigo):
                    continue

                descricao = cells[idx_desc] if idx_desc < len(cells) else ""
                qtd = _limpar_quantidade(cells[idx_qtd]) if idx_qtd < len(cells) else 0

                dados.append({
                    "Código": codigo,
                    "Descrição": _limpar_descricao(descricao),
                    "Estoque_Atual": qtd
                })

    except Exception as e:
        logger.error(f"Erro ao ler DOCX: {e}")
        return pd.DataFrame(columns=["Código", "Descrição", "Estoque_Atual"])

    if not dados:
        return pd.DataFrame(columns=["Código", "Descrição", "Estoque_Atual"])

    df = pd.DataFrame(dados)
    df = df.groupby("Código", as_index=False).agg({
        "Descrição": "first",
        "Estoque_Atual": "sum"
    })
    return df


# ══════════════════════════════════════════════════════════════
# INGESTÃO DE PEDIDOS — PDF (Onfinity)
# ══════════════════════════════════════════════════════════════

def _is_continuation_page_text(text: str) -> bool:
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
    has_mapa = bool(re.search(r'Mapa\s+de\s+Separa', text, re.IGNORECASE))
    has_pedido_header = bool(re.search(r'Pedido\s+Cliente', text))
    has_cnpj = bool(re.search(r'CNPJ/CPF', text))
    has_vendedor = bool(re.search(r'Vendedor', text))

    return has_mapa and not (has_pedido_header and has_cnpj and has_vendedor)


def extrair_pedidos_pdf(arquivos) -> pd.DataFrame:
    """
    Extrai itens de pedidos do PDF Onfinity (Mapa de Separação por Pedido).

    ESTRATÉGIA: Usa extração por TEXTO (linhas), não por tabela, porque o
    pdfplumber mescla todas as colunas numa única célula neste layout.

    Suporta pedidos multi-página: quando uma página não possui cabeçalho
    de identificação (Pedido/Cliente/CNPJ), seus itens são atribuídos
    ao pedido da página anterior.

    Formato de cada item no texto (pode ocupar 2-3 linhas):
      [Local] CÓDIGO [CodFornecedor] DESC... UNID QTD[,00] [PESO] [$ PREÇO]
      [continuação da descrição]

    Exemplo real:
      R01.B01.A3 6250200002 843334 1231-CANETA CANETA UN 10,00 0,18 $ 6,86
      ESFEROGRAFICA BIC 4 CORES
    """
    itens = []
    continuation_count = 0

    # Regex para linhas de item: captura Código (7-10 dígitos) seguido de dados
    # Padrão: [opcional_local] CODIGO_SKU [codigo_fornecedor] DESCRICAO UNID QTD...
    #
    # IMPORTANTE: Usamos regex GREEDY (.+) para a descrição, fazendo o regex avançar
    # o máximo possível e depois fazer backtrack para encontrar o ÚLTIMO par
    # UNIDADE+QUANTIDADE na linha. Isso é crítico porque descrições como
    # "SACO LIXO 200L PCT 100UN" contêm unidades embutidas no texto.
    # Com lazy (.+?), a regex capturava "PCT 100" da descrição em vez do
    # "PCT 2,00" real da coluna Quant.
    RE_ITEM_LINE = re.compile(
        r'(?:^|\s)(\d{7,10})\s+'   # Código SKU (7-10 dígitos)
        r'(.+)'                      # Descrição (GREEDY — avança até o último par unid+qtd)
        r'\b(UN|PCT|CX|RL|KG|MT|PR|DZ|CT|MH|LT|GL|SC|BD|FD|RS|TB|JG|PC|PAR|BL)\b'  # Unidade
        r'\s+(\d{1,6}[,.]?\d{0,2})'  # Quantidade
    )

    # Regex para capturar o campo Local (ex: R01.B07.A3, R02.B10.A1)
    RE_LOCAL = re.compile(r'\b(R\d{2}\.B\d{2}\.A\d{1,2})\b')

    for arquivo in arquivos:
        arquivo.seek(0)
        last_pedido = None   # Último pedido válido (para herança em continuações)
        last_cliente = None  # Último cliente válido

        with pdfplumber.open(arquivo) as pdf:
            for page_idx, page in enumerate(pdf.pages):
                text = page.extract_text() or ""
                if not text.strip():
                    continue

                lines = text.split('\n')

                # --- Detecção de continuação ---
                is_continuation = _is_continuation_page_text(text)

                if is_continuation:
                    continuation_count += 1
                    if last_pedido is not None:
                        pedido = last_pedido
                        cliente = last_cliente
                        logger.info(
                            f"Checklist: Página {page_idx+1} detectada como "
                            f"CONTINUAÇÃO do Pedido {pedido}"
                        )
                    else:
                        # Continuação órfã — sem pedido anterior
                        logger.error(
                            f"Checklist: Página {page_idx+1} é continuação, "
                            f"mas sem pedido anterior. Itens descartados."
                        )
                        continue
                else:
                    # --- Cabeçalho: Pedido e Cliente ---
                    pedido = "Desconhecido"
                    cliente = "Desconhecido"

                    m_ped = re.search(r'Pedido\s+Cliente\s*\n\s*(\d{4,10})', text)
                    if m_ped:
                        pedido = m_ped.group(1).strip()
                    else:
                        m_ped2 = re.search(r'Pedido[^\d]*(\d{4,10})', text, re.IGNORECASE)
                        if m_ped2:
                            pedido = m_ped2.group(1).strip()

                    m_cli = re.search(
                        r'Pedido\s+Cliente\s*\n\s*\d+\s+\d*\s*[-–]?\s*(.+?)(?:\s+CNPJ|$)',
                        text
                    )
                    if m_cli:
                        cliente = m_cli.group(1).strip()
                        cliente = re.sub(r'^\d+\s*[-–]\s*', '', cliente).strip()

                    # Atualizar referência para futuras continuações
                    last_pedido = pedido
                    last_cliente = cliente

                # --- Extrair itens linha a linha ---
                # Mapeamento temporário de Local por código (pode vir em linhas diferentes)
                local_por_codigo = {}

                for line in lines:
                    line_stripped = line.strip()

                    # Ignorar linhas de cabeçalho/rodapé
                    if not line_stripped:
                        continue
                    upper = line_stripped.upper()
                    if any(kw in upper for kw in ['PEDIDO', 'CLIENTE', 'CNPJ', 'VENDEDOR',
                                                   'OBSERVA', 'RESP.', 'TOTAIS', 'MAPA DE',
                                                   'LOCAL', 'FORNECEDOR']):
                        if 'LOCAL' in upper and re.search(r'\d{7,10}', line_stripped):
                            pass  # É uma linha de item com Local preenchido
                        else:
                            continue

                    # Capturar Local (endereço físico no estoque) se presente na linha
                    m_local = RE_LOCAL.search(line_stripped)
                    current_local = m_local.group(1) if m_local else ""

                    # Associar Local ao código se a linha contém ambos
                    if current_local:
                        m_code_check = re.search(r'\d{7,10}', line_stripped)
                        if m_code_check:
                            local_por_codigo[m_code_check.group(0)] = current_local

                    # Com regex GREEDY (.+), o .search() já encontra o ÚLTIMO par UNID+QTD
                    m = RE_ITEM_LINE.search(line_stripped)
                    if not m:
                        continue

                    codigo = m.group(1).strip()
                    desc_raw = m.group(2).strip()
                    unidade = m.group(3).strip()
                    qtd_str = m.group(4).strip().replace(',', '.')

                    # Limpar descrição: remover código de fornecedor do início
                    descricao = _limpar_descricao(desc_raw)

                    # Converter quantidade
                    try:
                        qtd = int(float(qtd_str))
                    except (ValueError, TypeError):
                        qtd = 0

                    # Buscar Local associado a este código
                    local = local_por_codigo.get(codigo, current_local)

                    if qtd > 0:
                        itens.append({
                            "Pedido": pedido,
                            "Cliente": cliente,
                            "Código": codigo,
                            "Descrição": descricao,
                            "Qtd_Pedida": qtd,
                            "Local": local,
                        })

    if continuation_count > 0:
        logger.info(
            f"Checklist: ⚠️ {continuation_count} página(s) de continuação "
            f"detectada(s) e mesclada(s) com pedidos anteriores"
        )

    if not itens:
        return pd.DataFrame(columns=["Pedido", "Cliente", "Código", "Descrição", "Qtd_Pedida", "Local"])

    return pd.DataFrame(itens)


# ══════════════════════════════════════════════════════════════
# MOTOR DE CRUZAMENTO E ALOCAÇÃO
# ══════════════════════════════════════════════════════════════

def cruzar_dados(df_estoque: pd.DataFrame, df_pedidos: pd.DataFrame) -> dict:
    """
    Cruza estoque com pedidos e retorna dicionário com:
    - batch_df: DataFrame para Batch Picking (visão por produto)
    - pedidos_completos: lista de dicts {pedido, cliente, itens_df}
    - pedidos_parciais: lista de dicts {pedido, cliente, itens_df, faltas_df}
    - pedidos_sem_estoque: lista de dicts (100% ruptura)
    - stats: dicionário com estatísticas
    """
    # Normaliza códigos em ambos os DataFrames
    df_estoque = df_estoque.copy()
    df_pedidos = df_pedidos.copy()
    df_estoque["Código"] = df_estoque["Código"].astype(str).str.strip()
    df_pedidos["Código"] = df_pedidos["Código"].astype(str).str.strip()

    # Merge por Código (SKU) — chave primária
    # Traz Estoque_Atual e Descrição do estoque (mais completa que a do pedido)
    df_master = pd.merge(
        df_pedidos,
        df_estoque[["Código", "Estoque_Atual", "Descrição"]],
        on="Código",
        how="left",
        suffixes=("_Pedido", ""),
    )
    df_master["Estoque_Atual"] = df_master["Estoque_Atual"].fillna(0).astype(int)
    # Usa a descrição do estoque (completa com referência); fallback para a do pedido
    df_master["Descrição"] = df_master["Descrição"].fillna(df_master["Descrição_Pedido"])
    df_master.drop(columns=["Descrição_Pedido"], inplace=True)

    # ── BATCH PICKING ──
    def _agg_pedidos(group):
        """Agrega pedidos por SKU: 'Ped 1001 (10un), Ped 1002 (5un)'."""
        parts = []
        for _, r in group.iterrows():
            parts.append(f"Ped {r['Pedido']} ({int(r['Qtd_Pedida'])}un)")
        return ", ".join(parts)

    batch_df = df_master.groupby("Código", as_index=False).agg(
        Descrição=("Descrição", "first"),
        Qtd_Total=("Qtd_Pedida", "sum"),
        Estoque_Atual=("Estoque_Atual", "first"),
        Num_Pedidos=("Pedido", "nunique"),
        Local=("Local", lambda x: next((v for v in x if v), "")),  # Primeiro Local não-vazio
    )
    # Pedidos destino
    pedidos_destino = df_master.groupby("Código").apply(
        _agg_pedidos, include_groups=False
    ).reset_index()
    pedidos_destino.columns = ["Código", "Pedidos_Destino"]
    batch_df = batch_df.merge(pedidos_destino, on="Código", how="left")

    batch_df["Falta"] = (batch_df["Qtd_Total"] - batch_df["Estoque_Atual"]).clip(lower=0)
    batch_df["Status"] = batch_df["Falta"].apply(lambda x: "🔴 FALTA" if x > 0 else "🟢 OK")

    # ── DISCRETE PICKING (por pedido) ──
    pedidos_completos = []
    pedidos_parciais = []
    pedidos_sem_estoque = []

    for pedido_num in df_master["Pedido"].unique():
        df_ped = df_master[df_master["Pedido"] == pedido_num].copy()
        cliente = df_ped["Cliente"].iloc[0]
        df_ped["Falta"] = (df_ped["Qtd_Pedida"] - df_ped["Estoque_Atual"]).clip(lower=0)

        faltas = df_ped[df_ped["Falta"] > 0]
        total_itens = len(df_ped)
        itens_sem_estoque = len(df_ped[df_ped["Estoque_Atual"] == 0])

        info = {
            "pedido": pedido_num,
            "cliente": cliente,
            "itens_df": df_ped[["Local", "Código", "Descrição", "Qtd_Pedida", "Estoque_Atual", "Falta"]],
            "total_itens": total_itens,
            "itens_ok": total_itens - len(faltas),
        }

        if faltas.empty:
            pedidos_completos.append(info)
        elif itens_sem_estoque == total_itens:
            info["faltas_df"] = faltas
            pedidos_sem_estoque.append(info)
        else:
            info["faltas_df"] = faltas
            pedidos_parciais.append(info)

    stats = {
        "total_skus": len(batch_df),
        "skus_ok": len(batch_df[batch_df["Falta"] == 0]),
        "skus_falta": len(batch_df[batch_df["Falta"] > 0]),
        "total_pedidos": df_master["Pedido"].nunique(),
        "pedidos_completos": len(pedidos_completos),
        "pedidos_parciais": len(pedidos_parciais),
        "pedidos_sem_estoque": len(pedidos_sem_estoque),
        "total_itens_pedidos": int(df_master["Qtd_Pedida"].sum()),
    }

    return {
        "batch_df": batch_df,
        "master_df": df_master,
        "pedidos_completos": pedidos_completos,
        "pedidos_parciais": pedidos_parciais,
        "pedidos_sem_estoque": pedidos_sem_estoque,
        "stats": stats,
    }
