"""
danfe_parser.py — Motor de extração de dados de DANFE (Nota Fiscal Eletrônica).

Responsável por:
- Leitura de PDFs de Notas Fiscais (DANFE) via pdfplumber
- Extração de campos-chave usando regex agressivo
- Normalização de valores (datas, pesos, valores monetários BR)
- Geração de DataFrame estruturado para exportação Excel

Colunas de saída (ordem estrita):
['MES', 'Data', 'Veiculo', 'Operacao', 'Remetente', 'Cliente',
 'Bairro', 'UF', 'Nota_Fiscal', 'Pedido', 'Peso', 'Volumes', 'Valor_Nota']
"""

import re
import io
import logging
from typing import Optional

import pandas as pd
import pdfplumber

logger = logging.getLogger(__name__)

# ══════════════════════════════════════════════════════════════
# CONSTANTES
# ══════════════════════════════════════════════════════════════

COLUNAS_SAIDA = [
    'MES', 'Data', 'Veiculo', 'Operacao', 'Remetente', 'Cliente',
    'Bairro', 'UF', 'Nota_Fiscal', 'Pedido', 'Peso', 'Volumes', 'Valor_Nota'
]

MES_SIGLAS = {
    '01': 'JAN', '02': 'FEV', '03': 'MAR', '04': 'ABR',
    '05': 'MAI', '06': 'JUN', '07': 'JUL', '08': 'AGO',
    '09': 'SET', '10': 'OUT', '11': 'NOV', '12': 'DEZ',
}

MES_LOWER = {
    '01': 'jan', '02': 'fev', '03': 'mar', '04': 'abr',
    '05': 'mai', '06': 'jun', '07': 'jul', '08': 'ago',
    '09': 'set', '10': 'out', '11': 'nov', '12': 'dez',
}

# Remetentes conhecidos — palavras-chave no texto da DANFE
REMETENTES_CONHECIDOS = [
    "ONFINITY",
    "GREEN BAGS",
]

# ══════════════════════════════════════════════════════════════
# HELPERS DE NORMALIZAÇÃO
# ══════════════════════════════════════════════════════════════

def _extrair_nota_fiscal(texto: str) -> Optional[str]:
    """
    Extrai número da nota fiscal.
    Busca padrões como 'Nº 000.002.426', 'N.: 000.002.426', 'NF-e Nº 2426'
    Remove pontos, zeros à esquerda. Retorna como string limpa.
    """
    # Padrão principal: Nº/N.:/N° seguido de número com pontos (ex: 000.002.426)
    m = re.search(
        r'N[ºo°.:]+\s*:?\s*([\d.]+)',
        texto,
        re.IGNORECASE
    )
    if m:
        numero = m.group(1).replace('.', '')
        # Remove zeros à esquerda
        return str(int(numero)) if numero.isdigit() else numero

    # Fallback: número de NF direto após "NOTA FISCAL"
    m2 = re.search(r'NOTA\s+FISCAL[^\d]*([\d.]+)', texto, re.IGNORECASE)
    if m2:
        numero = m2.group(1).replace('.', '')
        return str(int(numero)) if numero.isdigit() else numero

    return None


def _extrair_data_emissao(texto: str) -> tuple[Optional[str], Optional[str]]:
    """
    Extrai data de emissão e retorna (MES_sigla, Data_formatada).

    Busca formatos: dd/mm/aaaa ou dd/mm/aa
    Retorna: ('MAI', '19/mai') ou (None, None)
    """
    # DATA DE EMISSÃO / DATA DA EMISSÃO / EMISSÃO
    m = re.search(
        r'(?:DATA\s+D[AE]\s+EMISS[ÃA]O|EMISS[ÃA]O)\s*[:\s]*(\d{2})[/\-](\d{2})[/\-](\d{2,4})',
        texto,
        re.IGNORECASE
    )
    if m:
        dia, mes, ano = m.group(1), m.group(2), m.group(3)
        mes_sigla = MES_SIGLAS.get(mes, mes)
        mes_lower = MES_LOWER.get(mes, mes)
        data_fmt = f"{dia}/{mes_lower}"
        return mes_sigla, data_fmt

    # Fallback genérico: primeira data dd/mm/aaaa no texto
    m2 = re.search(r'(\d{2})[/\-](\d{2})[/\-](\d{2,4})', texto)
    if m2:
        dia, mes = m2.group(1), m2.group(2)
        mes_sigla = MES_SIGLAS.get(mes, mes)
        mes_lower = MES_LOWER.get(mes, mes)
        data_fmt = f"{dia}/{mes_lower}"
        return mes_sigla, data_fmt

    return None, None


def _extrair_remetente(texto: str) -> str:
    """
    Identifica o remetente (emitente) buscando por palavras-chave
    conhecidas nas primeiras linhas do texto.
    """
    # Procura por remetentes conhecidos
    texto_upper = texto.upper()
    for remetente in REMETENTES_CONHECIDOS:
        if remetente in texto_upper:
            return remetente

    # Fallback: tenta capturar o nome no bloco RAZÃO SOCIAL do emitente
    # Geralmente aparece logo abaixo de "DANFE" ou próximo ao logo
    linhas = texto.split('\n')
    for linha in linhas[:15]:  # Primeiras 15 linhas
        linha_limpa = linha.strip()
        # Procura por nome de empresa (palavra com LTDA, S.A, S/A, ME, EIRELI, etc.)
        if re.search(r'\b(LTDA|S\.?A\.?|S/A|ME|EIRELI|EPP|COMERCIAL)\b', linha_limpa, re.IGNORECASE):
            # Limpa e retorna o nome
            nome = re.sub(r'\s+', ' ', linha_limpa).strip()
            # Remove prefixos de campo comuns
            nome = re.sub(r'^(?:RAZ[ÃA]O\s+SOCIAL|NOME)\s*[:/]?\s*', '', nome, flags=re.IGNORECASE)
            if len(nome) > 3:
                return nome

    return ""


def _extrair_cliente(texto: str) -> str:
    """
    Extrai a Razão Social do destinatário.

    No pdfplumber real, o layout é:
      Linha header: "NOME / RAZÃO SOCIAL CNPJ / CPF DATA DA EMISSÃO"
      Linha dados:  "BRF S.A. 01.838.723/0308-91 19/05/26"

    Estratégia: busca o bloco DESTINATÁRIO, pula a linha de header,
    captura a linha de dados e limpa CNPJ/datas dela.
    """
    # Padrão 1: bloco DESTINATÁRIO → header NOME → linha de dados
    m = re.search(
        r'DESTINAT[ÁA]RIO[^\n]*\n'
        r'[^\n]*(?:NOME|RAZ[ÃA]O)[^\n]*\n'
        r'\s*([^\n]+)',
        texto,
        re.IGNORECASE
    )
    if m:
        cliente = _limpar_campo_cliente(m.group(1))
        if cliente:
            return cliente

    # Padrão 2: "INFORMAÇÕES DO LOCAL DE ENTREGA" → NOME → dados
    m2 = re.search(
        r'LOCAL\s+DE\s+ENTREGA[^\n]*\n'
        r'[^\n]*(?:NOME|RAZ[ÃA]O)[^\n]*\n'
        r'\s*([^\n]+)',
        texto,
        re.IGNORECASE
    )
    if m2:
        cliente = _limpar_campo_cliente(m2.group(1))
        if cliente:
            return cliente

    # Padrão 3: Fallback na linha "Dest/Rem:" do cabeçalho de recebimento
    m3 = re.search(
        r'Dest/Rem:\s*([^|\n]+)',
        texto,
        re.IGNORECASE
    )
    if m3:
        cliente = m3.group(1).strip()
        # Remove sufixos como "| Fantasia: ..."
        cliente = re.split(r'\s*\|', cliente)[0].strip()
        if len(cliente) > 2:
            return cliente

    return ""


def _limpar_campo_cliente(raw: str) -> str:
    """Limpa a linha de dados do cliente, removendo CNPJ, datas e códigos."""
    cliente = raw.strip()
    # Corta antes do primeiro CNPJ (dd.ddd.ddd/dddd-dd) ou CPF ou data
    cliente = re.split(
        r'\s+\d{2}\.\d{3}\.\d{3}[/\\]\d{4}[-]\d{2}',  # CNPJ
        cliente
    )[0]
    # Corta também antes de sequências numéricas longas (CPF, datas)
    cliente = re.split(r'\s+\d{2,3}\.\d{3}', cliente)[0]
    # Remove campos remanescentes
    cliente = re.split(
        r'\s+(?:CNPJ|CPF|DATA|INSCRI)',
        cliente,
        flags=re.IGNORECASE
    )[0]
    cliente = cliente.strip()
    return cliente if len(cliente) > 2 else ""


def _extrair_peso_bruto(texto: str) -> str:
    """
    Extrai o Peso Bruto.

    No pdfplumber real, o layout é:
      Header: "QUANTIDADE ESPÉCIE MARCA NUMERAÇÃO PESO BRUTO PESO LÍQUIDO"
      Dados:  "10 VOLUMES 41,250 41,250"
    Ou no layout com newline:
      "PESO BRUTO\n18.750,55"

    Estratégia: busca o header da linha transportadora, captura os valores
    numéricos da linha de dados. O penúltimo valor numérico é PESO BRUTO.
    """
    # Padrão 1: Header com PESO BRUTO seguido de dados na próxima linha
    # Captura a linha inteira de dados após o header de QUANTIDADE...PESO BRUTO
    m = re.search(
        r'QUANTIDADE\s+ESP[ÉE]CIE.*?PESO\s+BRUTO.*?PESO\s+L[ÍI]QUIDO\s*\n'
        r'\s*([^\n]+)',
        texto,
        re.IGNORECASE
    )
    if m:
        linha_dados = m.group(1).strip()
        # Extrair todos os valores numéricos da linha
        numeros = re.findall(r'[\d.,]+', linha_dados)
        # Formato: QTD [ESPECIE] [MARCA] [NUMERACAO] PESO_BRUTO PESO_LIQ
        # O penúltimo número é PESO BRUTO
        if len(numeros) >= 2:
            return _normalizar_valor_br(numeros[-2])
        elif len(numeros) == 1:
            return _normalizar_valor_br(numeros[0])

    # Padrão 2: PESO BRUTO com valor na mesma linha ou próxima (layout vertical)
    m2 = re.search(
        r'PESO\s+BRUTO\s*[:\s]*\n?\s*([\d.,]+)',
        texto,
        re.IGNORECASE
    )
    if m2:
        return _normalizar_valor_br(m2.group(1).strip())

    # Padrão 3: PESO BRUTO genérico
    m3 = re.search(
        r'PESO\s+BRUTO[^\d]*([\d.,]+)',
        texto,
        re.IGNORECASE
    )
    if m3:
        return _normalizar_valor_br(m3.group(1).strip())

    return ""


def _extrair_volumes(texto: str) -> str:
    """
    Extrai a quantidade de volumes.

    No pdfplumber real, o layout é:
      Header: "QUANTIDADE ESPÉCIE MARCA NUMERAÇÃO PESO BRUTO PESO LÍQUIDO"
      Dados:  "10 VOLUMES 41,250 41,250"
    Ou layout vertical:
      "QUANTIDADE\n6"

    O primeiro número da linha de dados é a QUANTIDADE.
    """
    # Padrão 1: header da tabela de transporte, dados na linha seguinte
    m = re.search(
        r'QUANTIDADE\s+ESP[ÉE]CIE.*?PESO\s+(?:BRUTO|L[ÍI]QUIDO)\s*\n'
        r'\s*(\d+)',
        texto,
        re.IGNORECASE
    )
    if m:
        return m.group(1).strip()

    # Padrão 2: QUANTIDADE com valor vertical (newline)
    m2 = re.search(
        r'QUANTIDADE\s*\n\s*(\d+)',
        texto,
        re.IGNORECASE
    )
    if m2:
        return m2.group(1).strip()

    # Padrão 3: QUANTIDADE com valor na mesma linha
    m3 = re.search(
        r'QUANTIDADE[^\d\n]*(\d+)',
        texto,
        re.IGNORECASE
    )
    if m3:
        return m3.group(1).strip()

    return ""


def _extrair_valor_nota(texto: str) -> str:
    """
    Extrai o Valor Total da Nota.

    No pdfplumber real, o layout é multi-linha:
      Header: "VALOR DO FRETE VALOR DO SEGURO DESCONTO ... VALOR TOTAL DA NOTA"
      Dados:  "0,00 0,00 0,00 0,00 0,00 1.102,50"

    Estratégia: busca "VALOR TOTAL DA NOTA" no header, e o ÚLTIMO número
    da linha de dados seguinte é o Valor Total da Nota.
    """
    # Padrão 1: VALOR TOTAL DA NOTA no final do header, dados na próxima linha
    m = re.search(
        r'VALOR\s+TOTAL\s+DA\s+NOTA\s*\n\s*([^\n]+)',
        texto,
        re.IGNORECASE
    )
    if m:
        linha_dados = m.group(1).strip()
        # O último número da linha é o Valor Total da Nota
        numeros = re.findall(r'[\d.,]+', linha_dados)
        if numeros:
            return _normalizar_valor_br(numeros[-1])

    # Padrão 2: valor na mesma linha (layout vertical)
    m2 = re.search(
        r'VALOR\s+TOTAL\s+DA\s+NOTA\s*[:\s]*([\d.,]+)',
        texto,
        re.IGNORECASE
    )
    if m2:
        return _normalizar_valor_br(m2.group(1).strip())

    # Padrão 3: genérico — qualquer número após "VALOR TOTAL DA NOTA"
    m3 = re.search(
        r'VALOR\s+TOTAL\s+DA\s+NOTA[^\d]*([\d.,]+)',
        texto,
        re.IGNORECASE
    )
    if m3:
        return _normalizar_valor_br(m3.group(1).strip())

    return ""


def _normalizar_valor_br(valor: str) -> str:
    """
    Normaliza valor monetário/peso no formato BR.
    Remove ponto de milhar, mantém vírgula decimal.
    Ex: '1.547,50' → '1547,50'
    Ex: '613,50' → '613,50'
    Ex: '18750' → '18750'
    """
    if not valor:
        return ""

    # Se tem vírgula → formato BR (ponto é milhar)
    if ',' in valor:
        # Remove pontos de milhar
        return valor.replace('.', '')

    # Se tem ponto → pode ser decimal ou milhar
    # Heurística: se tem 1-2 dígitos após o último ponto → decimal
    parts = valor.rsplit('.', 1)
    if len(parts) == 2 and len(parts[1]) <= 2:
        # É decimal → converte para formato BR
        return valor.replace('.', ',')

    # Caso contrário, é milhar
    return valor.replace('.', '')


def _extrair_bairro(texto: str, cliente: str) -> str:
    """
    Determina o Bairro (cidade de entrega) com base em regras de negócio.

    Regra atual:
    - Aplica-se SOMENTE quando o Cliente começa com 'BRF' ou 'MOGIANA'.
    - Busca no bloco 'DADOS ADICIONAIS' / 'INFORMAÇÕES COMPLEMENTARES'
      as palavras-chave 'EMBU' ou 'TABOAO' (acentuado ou não).
    - Se encontrar APENAS 'EMBU'  → retorna 'EMBU DAS ARTES'
    - Se encontrar APENAS 'TABOAO' → retorna 'TABOAO DA SERRA'
    - Se encontrar AMBAS ou NENHUMA → retorna '' (vazio)

    Args:
        texto: texto completo da página extraído pelo pdfplumber
        cliente: nome do cliente já extraído

    Returns:
        string com o bairro/cidade ou vazio
    """
    # Só aplica a regra para clientes BRF ou MOGIANA
    cliente_upper = (cliente or "").strip().upper()
    if not (cliente_upper.startswith("BRF") or cliente_upper.startswith("MOGIANA")):
        return ""

    # Isola o bloco DADOS ADICIONAIS (rodapé)
    bloco_adicional = ""
    m = re.search(
        r'DADOS\s+ADICIONAIS(.+)',
        texto,
        re.IGNORECASE | re.DOTALL
    )
    if m:
        bloco_adicional = m.group(1).upper()
    else:
        # Fallback: busca INFORMAÇÕES COMPLEMENTARES
        m2 = re.search(
            r'INFORMA[ÇC][ÕO]ES\s+COMPLEMENTARES(.+)',
            texto,
            re.IGNORECASE | re.DOTALL
        )
        if m2:
            bloco_adicional = m2.group(1).upper()

    if not bloco_adicional:
        return ""

    # Busca as palavras-chave (aceita variações: EMBU, TABOAO, TABOÃO)
    tem_embu = bool(re.search(r'\bEMBU\b', bloco_adicional))
    tem_taboao = bool(re.search(r'\bTABOA[OÃ]', bloco_adicional))

    if tem_embu and tem_taboao:
        # Ambas encontradas → ambíguo, deixa vazio
        return ""
    elif tem_embu:
        return "EMBU DAS ARTES"
    elif tem_taboao:
        return "TABOAO DA SERRA"

    return ""


def _extrair_pedido_associado(texto: str) -> str:
    """
    Busca o número do pedido no rodapé da DANFE (DADOS ADICIONAIS).

    Padrões reconhecidos (case-insensitive):
      - 'Pedido(s): 099613'
      - 'Pedido: 2771'
      - 'Nº Pedido: 2771'
      - 'Nr. Pedido: 2771'
      - 'Nº do Pedido: 2771'
      - 'PED: 2771' / 'PED 2771'
      - 'Pedido 2771' (sem dois-pontos)
      - 'Pedido nº 2771'
    """
    # Lista de padrões por prioridade (mais específico primeiro)
    padroes = [
        # 'Pedido(s): NNNNN' ou 'Pedido(s) NNNNN'
        r'Pedido(?:s)?\s*:?\s*(?:n[ºo°]?\.?\s*)?(\d{3,7})',
        # 'Nº Pedido: NNNNN' / 'Nr. Pedido: NNNNN' / 'Nº do Pedido: NNNNN'
        r'N[ºo°r]\.?\s*(?:do\s+)?Pedido\s*:?\s*(\d{3,7})',
        # 'PED: NNNNN' ou 'PED NNNNN'
        r'\bPED\s*:?\s*(\d{3,7})',
        # 'Pedido compra: NNNNN' / 'Pedido de compra: NNNNN'
        r'Pedido\s+(?:de\s+)?compra\s*:?\s*(\d{3,7})',
    ]

    for padrao in padroes:
        m = re.search(padrao, texto, re.IGNORECASE)
        if m:
            return m.group(1)

    return ""


# ══════════════════════════════════════════════════════════════
# FUNÇÃO PRINCIPAL DE EXTRAÇÃO
# ══════════════════════════════════════════════════════════════

def extrair_danfe(arquivo) -> dict:
    """
    Extrai dados de um único PDF DANFE.

    Estratégia: Extrai texto de toda a primeira página usando
    page.extract_text() e aplica blocos de re.search() agressivos.
    Trata None caso um padrão falhe para evitar quebra.

    Args:
        arquivo: file-like object (UploadedFile do Streamlit ou file handle)

    Returns:
        dict com as chaves correspondentes às colunas de saída,
        ou dict vazio se a extração falhar completamente.
    """
    try:
        arquivo.seek(0)
    except (AttributeError, io.UnsupportedOperation):
        pass

    texto_completo = ""

    try:
        with pdfplumber.open(arquivo) as pdf:
            if not pdf.pages:
                logger.warning("PDF sem páginas.")
                return {}

            # Extrai texto da primeira página (DANFE principal)
            page = pdf.pages[0]
            texto_completo = page.extract_text() or ""

    except Exception as e:
        logger.error(f"Erro ao abrir PDF: {e}")
        return {}

    if not texto_completo.strip():
        logger.warning("Texto extraído está vazio.")
        return {}

    # --- Extração de cada campo com fallback seguro ---
    nota_fiscal = _extrair_nota_fiscal(texto_completo) or ""
    mes_sigla, data_fmt = _extrair_data_emissao(texto_completo)
    remetente = _extrair_remetente(texto_completo)
    cliente = _extrair_cliente(texto_completo)
    peso = _extrair_peso_bruto(texto_completo)
    volumes = _extrair_volumes(texto_completo)
    valor_nota = _extrair_valor_nota(texto_completo)
    bairro = _extrair_bairro(texto_completo, cliente)
    pedido_associado = _extrair_pedido_associado(texto_completo)

    return {
        'MES': mes_sigla or "",
        'Data': data_fmt or "",
        'Veiculo': "",         # Preenchido futuramente
        'Operacao': "",        # Preenchido futuramente
        'Remetente': remetente,
        'Cliente': cliente,
        'Bairro': bairro,
        'UF': "SP",            # Valor estático
        'Nota_Fiscal': nota_fiscal,
        'Pedido': pedido_associado,
        'Peso': peso,
        'Volumes': volumes,
        'Valor_Nota': valor_nota,
    }


def extrair_multiplos_danfe(arquivos: list) -> pd.DataFrame:
    """
    Processa múltiplos PDFs DANFE e retorna um DataFrame consolidado.

    Args:
        arquivos: lista de file-like objects (UploadedFiles do Streamlit)

    Returns:
        pd.DataFrame com colunas na ordem estrita de COLUNAS_SAIDA
    """
    registros = []

    for arquivo in arquivos:
        try:
            nome = getattr(arquivo, 'name', 'desconhecido')
            logger.info(f"Processando DANFE: {nome}")

            resultado = extrair_danfe(arquivo)
            if resultado:
                registros.append(resultado)
            else:
                logger.warning(f"Extração vazia para: {nome}")
                # Adiciona registro com campos vazios para manter rastreabilidade
                registros.append({col: "" for col in COLUNAS_SAIDA})

        except Exception as e:
            logger.error(f"Erro ao processar {getattr(arquivo, 'name', '?')}: {e}")
            registros.append({col: "" for col in COLUNAS_SAIDA})

    if not registros:
        return pd.DataFrame(columns=COLUNAS_SAIDA)

    df = pd.DataFrame(registros, columns=COLUNAS_SAIDA)
    return df


def gerar_excel_danfe(df: pd.DataFrame) -> bytes:
    """
    Gera bytes de um arquivo Excel (.xlsx) a partir do DataFrame de DANFEs.

    Args:
        df: DataFrame com as colunas de COLUNAS_SAIDA

    Returns:
        bytes do arquivo Excel pronto para download
    """
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Notas Fiscais")
    return buf.getvalue()
