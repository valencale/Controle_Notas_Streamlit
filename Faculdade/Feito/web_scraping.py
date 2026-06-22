"""
web_scraping.py — Script de Web Scraping para Portfólio de Web Analytics.
Disciplina: Web Analytics | Aluno: Alex Valencia | 2026

Extrai títulos, preços e avaliações de livros do site books.toscrape.com
usando BeautifulSoup + Requests, e salva os resultados em CSV.
"""

import requests
from bs4 import BeautifulSoup
import csv
import os

# ══════════════════════════════════════════════════════════════
# CONFIGURAÇÃO
# ══════════════════════════════════════════════════════════════
URL = "http://books.toscrape.com/"
OUTPUT_DIR = os.path.dirname(os.path.abspath(__file__))
CSV_FILE = os.path.join(OUTPUT_DIR, "livros_scraping.csv")

# ══════════════════════════════════════════════════════════════
# FUNÇÕES
# ══════════════════════════════════════════════════════════════

def extrair_livros(url: str) -> list[dict]:
    """Faz requisição HTTP e extrai dados dos livros da página."""
    print(f"[>] Acessando: {url}")
    
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
    }
    
    response = requests.get(url, headers=headers)
    response.raise_for_status()
    print(f"[OK] Status HTTP: {response.status_code}")
    
    soup = BeautifulSoup(response.text, "html.parser")
    
    livros = []
    for article in soup.find_all("article", class_="product_pod"):
        titulo = article.h3.a["title"]
        preco = article.find("p", class_="price_color").text.strip()
        
        # Extrair avaliação (star rating)
        rating_tag = article.find("p", class_="star-rating")
        avaliacao = rating_tag["class"][1] if rating_tag else "N/A"
        
        # Verificar disponibilidade
        disponivel = article.find("p", class_="instock availability")
        status = "Em estoque" if disponivel else "Indisponível"
        
        livros.append({
            "titulo": titulo,
            "preco": preco,
            "avaliacao": avaliacao,
            "status": status
        })
    
    return livros


def salvar_csv(livros: list[dict], filepath: str):
    """Salva a lista de livros em um arquivo CSV."""
    with open(filepath, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=["titulo", "preco", "avaliacao", "status"])
        writer.writeheader()
        writer.writerows(livros)
    print(f"[SAVE] Arquivo salvo: {filepath}")


# ══════════════════════════════════════════════════════════════
# EXECUÇÃO PRINCIPAL
# ══════════════════════════════════════════════════════════════

if __name__ == "__main__":
    print("=" * 60)
    print("  WEB SCRAPING — Books to Scrape")
    print("  Disciplina: Web Analytics | Alex Valencia | 2026")
    print("=" * 60)
    print()
    
    # 1. Extrair dados
    livros = extrair_livros(URL)
    
    print(f"\n[OK] {len(livros)} livros extraidos com sucesso!\n")
    print("-" * 60)
    print(f"{'TÍTULO':<45} {'PREÇO':>8}")
    print("-" * 60)
    
    for livro in livros:
        titulo_curto = livro["titulo"][:42] + "..." if len(livro["titulo"]) > 42 else livro["titulo"]
        print(f"{titulo_curto:<45} {livro['preco']:>8}")
    
    print("-" * 60)
    
    # 2. Salvar em CSV
    print()
    salvar_csv(livros, CSV_FILE)
    
    print(f"\n[OK] Processo concluido! {len(livros)} registros exportados.")
    print(f"[FILE] Arquivo: {CSV_FILE}")
