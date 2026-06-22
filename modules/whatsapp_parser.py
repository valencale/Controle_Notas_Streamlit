"""
whatsapp_parser.py — Parser de mensagens WhatsApp de motoristas.

Extrai dados estruturados de viagem a partir de mensagens de texto
enviadas por motoristas no grupo de WhatsApp da logística.

PADRÕES SUPORTADOS (catalogados de 15+ mensagens reais):
    - Placas: RLK0E24, Kzg9d54, QJJ 9302, Rlk 0e24 HR
    - Pedidos: "17 Pedidos", "Pedido 57", "01 Pedido"
    - Entregas: "17 Entregas", "5 entregues", "Entregue 51"
    - Coletas: "01 coleta", "1 coleta realizada", "2 coleta realizada"
    - KM: "Km inicial 158320", "Km. Final. 158444", "Km final. 315966"
    - Ocorrências: "4 não deu tempo Sabará", "1 voltou braslimpo"
    - Motoristas: "José Soares", "José Soares/ Fernando"
    - Header chat export: "[15:12, 20/05/2026] Nome: ..."
"""

import re
from datetime import datetime
from typing import Optional

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from config import KNOWN_PLATES


# ══════════════════════════════════════════════════════════════
# REGEX PATTERNS (compiled for performance)
# ══════════════════════════════════════════════════════════════

# WhatsApp chat export header: [15:12, 20/05/2026] Nome Contato:
RE_CHAT_HEADER = re.compile(
    r"\[(\d{2}:\d{2}),\s*(\d{2}/\d{2}/\d{4})\]\s*([^:]+):\s*(.*)",
    re.DOTALL,
)

# Plate patterns — Brazilian formats (Mercosul + old format)
# Matches: QJJ9302, QJJ-9302, QJJ 9302, RLK0E24, Rlk 0e24, Kzg9d54
RE_PLATE = re.compile(
    r"\b([A-Za-z]{3})\s*[-]?\s*(\d[A-Za-z0-9]\d{2})\b",
)

# Pedidos: "17 Pedidos" or "Pedido 57" (case insensitive)
RE_PEDIDOS_BEFORE = re.compile(r"(\d+)\s*pedidos?", re.IGNORECASE)
RE_PEDIDOS_AFTER = re.compile(r"pedidos?\s*(\d+)", re.IGNORECASE)

# Entregas: "17 Entregas", "5 entregues", "Entregue 51"
RE_ENTREGAS_BEFORE = re.compile(r"(\d+)\s*(?:entregas?|entregues?)", re.IGNORECASE)
RE_ENTREGAS_AFTER = re.compile(r"(?:entregas?|entregues?)\s*(\d+)", re.IGNORECASE)

# Coletas: "01 coleta", "1 coleta realizada", "2 coleta realizada"
RE_COLETAS = re.compile(r"(\d+)\s*coletas?(?:\s+realizada)?", re.IGNORECASE)

# KM Inicial: "Km inicial 158320", "Km. inicial 277300"
RE_KM_INICIAL = re.compile(r"km\.?\s*inicial\s*(\d+)", re.IGNORECASE)

# KM Final: "Km. Final. 158444", "Km final 277444", "Km final. 315966"
RE_KM_FINAL = re.compile(r"km\.?\s*final\.?\s*(\d+)", re.IGNORECASE)

# Occurrence lines: start with a number, NOT followed by known field words
# e.g. "4 não deu tempo Sabará", "2 devolveu brf", "1 voltou braslimpo"
RE_OCCURRENCE = re.compile(
    r"^(\d+)\s+(?!pedido|entreg|coleta|km\b|caixa|fardo)(.+)",
    re.IGNORECASE | re.MULTILINE,
)

# Lines that are structural fields (to exclude from OBS)
RE_STRUCTURAL = re.compile(
    r"^\s*(?:\d+\s*(?:pedido|entreg|coleta)|"
    r"(?:pedido|entreg)\w*\s*\d+|"
    r"km\.?\s*(?:inicial|final)|"
    r"[A-Za-z]{3}\s*[-]?\s*\d[A-Za-z0-9]\d{2})",
    re.IGNORECASE,
)


# ══════════════════════════════════════════════════════════════
# PLATE NORMALIZATION
# ══════════════════════════════════════════════════════════════

def normalize_plate(raw_plate: str) -> Optional[str]:
    """
    Normaliza uma placa bruta para o formato canônico da frota.

    Tenta encontrar a placa no mapa KNOWN_PLATES.
    Se não encontrar, formata como ABC-1D23 e tenta novamente.

    Args:
        raw_plate: Texto bruto da placa (ex: "Rlk 0e24", "QJJ9302").

    Returns:
        Placa canônica (ex: "RLK-0E24") ou None se não reconhecida.
    """
    cleaned = raw_plate.strip()

    # Try direct lookup (case-sensitive first)
    if cleaned in KNOWN_PLATES:
        return KNOWN_PLATES[cleaned]

    # Try uppercase lookup
    upper = cleaned.upper().replace(" ", "").replace("-", "")
    for raw_key, canonical in KNOWN_PLATES.items():
        normalized_key = raw_key.upper().replace(" ", "").replace("-", "")
        if upper == normalized_key:
            return canonical

    # Format as ABC-1D23 and try
    if len(upper) == 7:
        formatted = f"{upper[:3]}-{upper[3:]}"
        if formatted in KNOWN_PLATES:
            return KNOWN_PLATES[formatted]

    return None


# ══════════════════════════════════════════════════════════════
# SINGLE MESSAGE PARSER
# ══════════════════════════════════════════════════════════════

def parse_single_message(text: str, default_date: Optional[datetime] = None) -> dict:
    """
    Extrai dados estruturados de uma única mensagem de motorista.

    Args:
        text: Texto bruto da mensagem (com ou sem header de chat export).
        default_date: Data padrão se não encontrada no header.

    Returns:
        Dict com campos: data, motorista, ajudante, placa, pedidos, entregas,
        coletas, km_inicial, km_final, km_rodado, ocorrencias, obs, texto_original,
        warnings (lista de alertas).
    """
    result = {
        "data": default_date or datetime.now(),
        "motorista": "",
        "ajudante": "",
        "placa": "",
        "pedidos": 0,
        "entregas": 0,
        "coletas": 0,
        "km_inicial": 0,
        "km_final": 0,
        "km_rodado": 0,
        "ocorrencias": [],
        "obs": "",
        "texto_original": text.strip(),
        "warnings": [],
    }

    body = text.strip()

    # ── 1. Extract date from chat header if present ──
    header_match = RE_CHAT_HEADER.match(body)
    if header_match:
        time_str, date_str, contact_name, message_body = header_match.groups()
        try:
            result["data"] = datetime.strptime(date_str.strip(), "%d/%m/%Y")
        except ValueError:
            pass
        body = message_body.strip()

    # ── 2. Split into lines for processing ──
    lines = [l.strip() for l in body.split("\n") if l.strip()]
    if not lines:
        return result

    # ── 3. Extract plates ──
    plates_found = []
    for line in lines:
        for m in RE_PLATE.finditer(line):
            raw = m.group(0)
            # Filter out suffixes like "HR" after plate
            canonical = normalize_plate(raw)
            if canonical:
                plates_found.append(canonical)

    if plates_found:
        result["placa"] = plates_found[0]  # Primary plate

    # ── 4. Extract driver name ──
    # The driver name is typically the first line that is NOT a plate
    # and NOT a structural field
    for line in lines:
        stripped = line.strip()
        # Skip empty
        if not stripped:
            continue
        # Skip if it's a plate line
        plate_in_line = RE_PLATE.search(stripped)
        if plate_in_line:
            canonical = normalize_plate(plate_in_line.group(0))
            if canonical:
                continue
        # Skip structural fields
        if RE_STRUCTURAL.match(stripped):
            continue
        # Skip if it's just a number
        if stripped.isdigit():
            continue

        # This is likely the driver name
        name = stripped
        # Remove emojis
        name = re.sub(r'[^\w\s/.À-ÿ-]', '', name, flags=re.UNICODE).strip()
        # Check for ajudante pattern: "José Soares/ Fernando"
        if "/" in name:
            parts = [p.strip() for p in name.split("/", 1)]
            result["motorista"] = parts[0]
            result["ajudante"] = parts[1] if len(parts) > 1 else ""
        else:
            result["motorista"] = name
        break

    # ── 5. Extract Pedidos ──
    for line in lines:
        m = RE_PEDIDOS_BEFORE.search(line)
        if m:
            result["pedidos"] = int(m.group(1))
            break
        m = RE_PEDIDOS_AFTER.search(line)
        if m:
            result["pedidos"] = int(m.group(1))
            break

    # ── 6. Extract Entregas ──
    for line in lines:
        m = RE_ENTREGAS_BEFORE.search(line)
        if m:
            result["entregas"] = int(m.group(1))
            break
        m = RE_ENTREGAS_AFTER.search(line)
        if m:
            result["entregas"] = int(m.group(1))
            break

    # ── 7. Extract Coletas ──
    total_coletas = 0
    for line in lines:
        m = RE_COLETAS.search(line)
        if m:
            total_coletas += int(m.group(1))
    result["coletas"] = total_coletas

    # ── 8. Extract KM ──
    m = RE_KM_INICIAL.search(body)
    if m:
        result["km_inicial"] = int(m.group(1))
    m = RE_KM_FINAL.search(body)
    if m:
        result["km_final"] = int(m.group(1))

    # Calculate KM rodado
    if result["km_inicial"] > 0 and result["km_final"] > 0:
        result["km_rodado"] = result["km_final"] - result["km_inicial"]
        if result["km_rodado"] < 0:
            result["warnings"].append(
                f"KM Final ({result['km_final']}) menor que KM Inicial "
                f"({result['km_inicial']}). Possível erro de digitação."
            )

    # ── 9. Extract Occurrences ──
    occurrences = []
    for line in lines:
        m = RE_OCCURRENCE.match(line.strip())
        if m:
            qty = int(m.group(1))
            reason = m.group(2).strip()
            # Skip lines that are actually coleta/carga details
            if re.search(r"coleta|caixa|fardo|volume", reason, re.IGNORECASE):
                continue
            occurrences.append({"qtd": qty, "motivo": reason})
    result["ocorrencias"] = occurrences

    # ── 10. Extract OBS (remaining free text) ──
    obs_lines = []
    # Track which lines were consumed by structured fields
    used_patterns = [
        RE_PEDIDOS_BEFORE, RE_PEDIDOS_AFTER,
        RE_ENTREGAS_BEFORE, RE_ENTREGAS_AFTER,
        RE_COLETAS, RE_KM_INICIAL, RE_KM_FINAL,
    ]

    for line in lines:
        stripped = line.strip()
        if not stripped:
            continue
        # Skip the driver name line
        if stripped == result.get("_name_line", ""):
            continue
        # Skip plate lines
        plate_m = RE_PLATE.search(stripped)
        if plate_m and normalize_plate(plate_m.group(0)):
            continue
        # Skip if matched by any structural regex
        is_structural = False
        for pat in used_patterns:
            if pat.search(stripped):
                is_structural = True
                break
        if is_structural:
            continue
        # Skip if it's the driver name
        name_clean = re.sub(r'[^\w\s/.À-ÿ-]', '', stripped, flags=re.UNICODE).strip()
        if name_clean and (name_clean == result["motorista"] or
                          name_clean == f"{result['motorista']}/ {result['ajudante']}".strip("/ ")):
            continue
        # Skip occurrence lines already captured
        occ_match = RE_OCCURRENCE.match(stripped)
        if occ_match and not re.search(r"coleta|caixa|fardo|volume", occ_match.group(2), re.IGNORECASE):
            continue

        obs_lines.append(stripped)

    result["obs"] = "\n".join(obs_lines)

    return result


# ══════════════════════════════════════════════════════════════
# BATCH PARSER — WhatsApp Chat Export (.txt)
# ══════════════════════════════════════════════════════════════

def parse_chat_export(text: str) -> list[dict]:
    """
    Processa um arquivo de exportação de chat do WhatsApp (.txt).

    Identifica cada mensagem pelo header [HH:MM, DD/MM/YYYY] e extrai
    os dados de viagem de cada uma que contenha padrões de relatório
    (pedidos, entregas, km).

    Args:
        text: Conteúdo completo do arquivo .txt exportado do WhatsApp.

    Returns:
        Lista de dicts (mesmo formato de parse_single_message),
        uma entrada por mensagem que contenha dados de viagem.
    """
    # Split by WhatsApp message headers
    # Pattern: [HH:MM, DD/MM/YYYY] Contact Name: message
    message_pattern = re.compile(
        r"(\[\d{2}:\d{2},\s*\d{2}/\d{2}/\d{4}\]\s*[^:]+:)",
    )

    # Split text preserving the headers
    parts = message_pattern.split(text)

    messages = []
    # Parts alternate: [pre-text, header1, body1, header2, body2, ...]
    i = 1
    while i < len(parts) - 1:
        header = parts[i].strip()
        body = parts[i + 1].strip()
        full_message = f"{header} {body}"
        i += 2

        # Only parse messages that look like trip reports
        # (must have at least KM or pedido/entrega info)
        has_km = RE_KM_INICIAL.search(full_message) or RE_KM_FINAL.search(full_message)
        has_orders = (RE_PEDIDOS_BEFORE.search(full_message) or
                     RE_PEDIDOS_AFTER.search(full_message))
        has_coleta = RE_COLETAS.search(full_message)

        if has_km or has_orders or has_coleta:
            parsed = parse_single_message(full_message)
            messages.append(parsed)

    return messages


# ══════════════════════════════════════════════════════════════
# FORMATTING HELPERS
# ══════════════════════════════════════════════════════════════

def format_ocorrencias(ocorrencias: list[dict]) -> str:
    """Formata lista de ocorrências para exibição/persistência."""
    if not ocorrencias:
        return ""
    return "; ".join(
        f"{o['qtd']}x {o['motivo']}" for o in ocorrencias
    )


def parse_ocorrencias_str(text: str) -> list[dict]:
    """Reconstrói lista de ocorrências a partir de string formatada."""
    if not text or not text.strip():
        return []
    result = []
    for part in text.split(";"):
        part = part.strip()
        m = re.match(r"(\d+)x\s*(.*)", part)
        if m:
            result.append({"qtd": int(m.group(1)), "motivo": m.group(2).strip()})
    return result
