"""
geocoder.py — Módulo de geocoding (conversão endereço → coordenadas).

Utiliza geopy com o provider Nominatim (gratuito, OpenStreetMap).
Implementa cache em memória para evitar requests repetidos e
rate limiting de 1 request/segundo conforme exigido pelo Nominatim.
"""

import time
from typing import Optional

from geopy.geocoders import Nominatim
from geopy.exc import GeocoderTimedOut, GeocoderServiceError

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from config import NOMINATIM_USER_AGENT, GEOCODE_TIMEOUT, MAP_CENTER

import streamlit as st

# Rate limiting control
_last_request_time: float = 0

@st.cache_data(show_spinner=False, ttl=3600*24*30) # Cache por 30 dias
def geocode_address(address: str) -> Optional[tuple[float, float]]:
    """
    Converte um endereço em coordenadas (latitude, longitude).

    Utiliza Nominatim (OpenStreetMap) com cache local e rate limiting.
    Se o geocoding falhar, retorna None.

    Args:
        address: Endereço completo (ex: "Avenida Otto Baumgart, 500, São Paulo, SP").

    Returns:
        Tuple (latitude, longitude) ou None se não encontrado.
    """
    global _last_request_time

    if not address or not address.strip():
        return None

    if not address or not address.strip():
        return None

    # Rate limiting (1 request per second for Nominatim)
    elapsed = time.time() - _last_request_time
    if elapsed < 1.0:
        time.sleep(1.0 - elapsed)

    try:
        geolocator = Nominatim(user_agent=NOMINATIM_USER_AGENT, timeout=GEOCODE_TIMEOUT)
        _last_request_time = time.time()

        # Try with full address first
        location = geolocator.geocode(address + ", Brasil")

        # If not found, try simplifying the address
        if location is None:
            # Remove complement, keep street + city
            simplified = _simplify_address(address)
            if simplified != address:
                time.sleep(1.0)  # Rate limit
                _last_request_time = time.time()
                location = geolocator.geocode(simplified + ", Brasil")

        if location:
            coords = (location.latitude, location.longitude)
            return coords

    except (GeocoderTimedOut, GeocoderServiceError) as e:
        print(f"Geocoding error for '{address}': {e}")

    return None


def geocode_addresses_batch(addresses: list[str]) -> dict[str, Optional[tuple[float, float]]]:
    """
    Geocodifica uma lista de endereços, retornando um dicionário.

    Args:
        addresses: Lista de endereços.

    Returns:
        Dicionário {endereço: (lat, lon) ou None}.
    """
    results = {}
    for addr in addresses:
        if addr and addr.strip():
            results[addr] = geocode_address(addr)
    return results


def _simplify_address(address: str) -> str:
    """
    Simplifica um endereço removendo complementos e detalhes irrelevantes
    para aumentar a chance de match no geocoding.

    Exemplo:
        "AVENIDA OTTO BAUMGART, 500 - LOJA 226 229 A - 2049900 - SAO PAULO/SP"
        →  "AVENIDA OTTO BAUMGART, 500, SAO PAULO, SP"
    """
    import re

    # Remove LOJA, SALA, BLOCO, ANDAR, etc.
    simplified = re.sub(r'\s*-\s*(?:LOJA|SALA|BLOCO|ANDAR|APT|APTO|CONJ|CONJUNTO)\s+[^-]*', '', address, flags=re.IGNORECASE)

    # Extract city/state from end (pattern: CITY/UF or CITY-UF)
    city_state = re.search(r'[-\s]+([\w\s]+)\s*/\s*(\w{2})\s*$', simplified)

    # Extract street and number
    street_number = re.match(r'^([^-]+?)(?:\s*-|$)', simplified)

    if street_number and city_state:
        return f"{street_number.group(1).strip()}, {city_state.group(1).strip()}, {city_state.group(2).strip()}"

    return simplified


def get_default_center() -> tuple[float, float]:
    """
    Retorna as coordenadas do centro padrão do mapa (Osasco, SP).

    Returns:
        Tuple (latitude, longitude).
    """
    return tuple(MAP_CENTER)
