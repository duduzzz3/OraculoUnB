"""
Scraper Arte Galeria – Milan Dusek
Coleta obras de https://artegaleria.com.br/milan-dusek/

Usa requests + BeautifulSoup (sem navegador).

Uso básico:
    python milan_dusek_artegalariaWS.py
    python milan_dusek_artegalariaWS.py --max-obras 60 --output saida.xlsx
"""
from __future__ import annotations

import argparse
import json
import os
import re
import sys
import time
from dataclasses import dataclass, field
from datetime import datetime
from typing import Any
from urllib.parse import urljoin

import pandas as pd
import requests
from bs4 import BeautifulSoup, Tag
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

# ─── Constantes ───────────────────────────────────────────────────────────────
BASE_URL    = "https://artegaleria.com.br"
ARTISTA_URL = f"{BASE_URL}/milan-dusek/"
ARTISTA_NOME = "Milan Dusek"
SEM_INFO    = "Sem informação"

EXCEL_COLUMNS = [
    "Título", "Autor", "Ano", "Técnica", "Dimensões",
    "Preço", "Descrição", "Link da obra", "Link da imagem da obra",
]

# Regex reutilizáveis
YEAR_RE  = re.compile(r"\b(19|20)\d{2}\b")
DIM_RE   = re.compile(
    r"\b(\d{1,4}(?:[.,]\d{1,2})?)\s*(?:x|×|X)\s*(\d{1,4}(?:[.,]\d{1,2})?)(?:\s*(?:x|×|X)\s*(\d{1,4}(?:[.,]\d{1,2})?))?(?:\s*cm\b)?",
    re.I,
)
PRICE_RE = re.compile(r"R\$\s*(\d{1,3}(?:[.\s]\d{3})*(?:,\d{2})?|\d+(?:,\d{2})?)")


# ─── Utilitários ──────────────────────────────────────────────────────────────

def limpar(valor: Any) -> str:
    if valor is None:
        return SEM_INFO
    t = str(valor).replace("\xa0", " ").replace("\u200b", "").strip()
    t = re.sub(r"\s+", " ", t)
    return t if t else SEM_INFO


def abs_url(href: str | None) -> str:
    if not href:
        return SEM_INFO
    href = href.strip()
    if href.startswith("//"):
        return "https:" + href
    return urljoin(BASE_URL, href)


def extrair_ano(texto: str) -> str:
    m = YEAR_RE.search(texto)
    if m:
        ano = int(m.group(0))
        if 1800 <= ano <= datetime.now().year + 1:
            return str(ano)
    return SEM_INFO


def padronizar_dimensoes(texto: str) -> str:
    if not texto or texto == SEM_INFO:
        return SEM_INFO
    m = DIM_RE.search(texto)
    if not m:
        return SEM_INFO

    def fmt(v: str | None) -> str | None:
        if not v:
            return None
        try:
            n = float(v.replace(",", "."))
            return str(int(n)) if n.is_integer() else str(round(n, 2)).rstrip("0").rstrip(".")
        except ValueError:
            return None

    partes = [fmt(m.group(1)), fmt(m.group(2))]
    if m.group(3):
        partes.append(fmt(m.group(3)))
    partes = [p for p in partes if p]
    return (" x ".join(partes) + " cm") if len(partes) >= 2 else SEM_INFO


def extrair_preco(texto: str) -> float | str:
    m = PRICE_RE.search(texto or "")
    if not m:
        return SEM_INFO
    raw = m.group(1).replace(".", "").replace(",", ".").replace(" ", "")
    try:
        return round(float(raw), 2)
    except ValueError:
        return SEM_INFO


def melhor_imagem(tag: Tag | None) -> str:
    """Escolhe a URL de maior resolução a partir de img srcset, data-src ou src."""
    if tag is None:
        return SEM_INFO
    img = tag.find("img")
    if img is None:
        return SEM_INFO
    srcset = img.get("srcset") or img.get("data-srcset") or ""
    if srcset:
        partes = [p.strip().split(" ")[0] for p in srcset.split(",") if p.strip()]
        if partes:
            return abs_url(partes[-1])
    for attr in ("data-src", "data-original", "src"):
        val = img.get(attr)
        if val and not val.startswith("data:"):
            return abs_url(val)
    return SEM_INFO


def make_session() -> requests.Session:
    retry = Retry(
        total=4, connect=4, read=4, status=4,
        backoff_factor=0.8,
        status_forcelist=(429, 500, 502, 503, 504),
        allowed_methods=("GET",),
        raise_on_status=False,
    )
    adapter = HTTPAdapter(max_retries=retry, pool_connections=8, pool_maxsize=8)
    session = requests.Session()
    session.mount("https://", adapter)
    session.mount("http://", adapter)
    session.headers.update({
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"
        ),
        "Accept": "text/html,application/xhtml+xml,*/*;q=0.9",
        "Accept-Language": "pt-BR,pt;q=0.9,en;q=0.7",
        "Cache-Control": "no-cache",
    })
    return session


def fetch(session: requests.Session, url: str, timeout: int = 30) -> str:
    resp = session.get(url, timeout=timeout)
    resp.raise_for_status()
    if not resp.encoding or resp.encoding.lower() in {"iso-8859-1", "ascii"}:
        resp.encoding = resp.apparent_encoding or "utf-8"
    return resp.text


# ─── Extração JSON-LD ─────────────────────────────────────────────────────────

def jsonld_from_soup(soup: BeautifulSoup) -> list[dict]:
    out: list[dict] = []
    for script in soup.select('script[type="application/ld+json"]'):
        txt = (script.string or "").strip()
        if not txt:
            continue
        try:
            obj = json.loads(txt)
            if isinstance(obj, list):
                out.extend(obj)
            else:
                out.append(obj)
        except Exception:
            pass
    return out


def jsonld_value(jsons: list[dict], *keys: str) -> str:
    for obj in jsons:
        if not isinstance(obj, dict):
            continue
        for k in keys:
            v = obj.get(k)
            if isinstance(v, str) and v.strip():
                return v.strip()
            if isinstance(v, dict):
                for sub in ("name", "url", "contentUrl"):
                    sv = v.get(sub)
                    if isinstance(sv, str) and sv.strip():
                        return sv.strip()
    return SEM_INFO


# ─── Descoberta de links ───────────────────────────────────────────────────────

def descobrir_links(session: requests.Session, url_inicial: str, max_paginas: int = 0) -> list[str]:
    """Percorre a listagem do artista e coleta todos os links de obras."""
    links: list[str] = []
    vistos: set[str] = set()
    url = url_inicial
    pagina = 1

    while True:
        if max_paginas and pagina > max_paginas:
            break

        print(f"Listagem página {pagina}: {url}")
        try:
            html = fetch(session, url)
        except Exception as e:
            print(f"  Erro ao buscar listagem: {e}", file=sys.stderr)
            break

        soup = BeautifulSoup(html, "lxml")

        # Coleta links de obras — tenta múltiplos padrões de CMS
        novos = 0
        for seletor in [
            # WooCommerce
            "ul.products li.product a.woocommerce-loop-product__link",
            "ul.products li.product h2.woocommerce-loop-product__title",
            ".product-grid .product-item a",
            # Genéricos gallery/portfolio
            "article.artwork a",
            ".obra-item a",
            ".gallery-item a",
            ".post-item a",
            "article a[href*='/obra']",
            "article a[href*='/produto']",
            # Qualquer âncora dentro de um elemento de produto
            "[class*='product'] a[href]",
            "[class*='artwork'] a[href]",
            ".entry-content a[href]",
        ]:
            encontrados = soup.select(seletor)
            if encontrados:
                for tag in encontrados:
                    href = tag.get("href", "")
                    if not href:
                        continue
                    full = abs_url(href)
                    # Filtra links que são do mesmo domínio e parecem ser obras únicas
                    if (
                        BASE_URL in full
                        and full not in vistos
                        and full != url_inicial.rstrip("/")
                        and not full.endswith(("/page/1/", "/?"))
                        and "category" not in full
                        and "tag" not in full
                        and "author" not in full
                    ):
                        vistos.add(full)
                        links.append(full)
                        novos += 1
                if novos:
                    break

        print(f"  {novos} links novos (total: {len(links)})")

        if not novos:
            print("  Nenhum link novo. Encerrando listagem.")
            break

        # Tenta avançar para próxima página
        prox = None
        for seletor_prox in [
            "a.next.page-numbers",
            ".nav-links a[aria-label='Next Page']",
            ".pagination a.next",
            "a[rel='next']",
            ".woocommerce-pagination a.next",
        ]:
            tag_prox = soup.select_one(seletor_prox)
            if tag_prox and tag_prox.get("href"):
                prox = abs_url(tag_prox["href"])
                break

        if not prox or prox == url:
            print("  Sem próxima página.")
            break

        url = prox
        pagina += 1
        time.sleep(0.5)

    return links


# ─── Extração da página interna ───────────────────────────────────────────────

@dataclass
class Obra:
    titulo:      str = SEM_INFO
    autor:       str = ARTISTA_NOME
    ano:         str = SEM_INFO
    tecnica:     str = SEM_INFO
    dimensoes:   str = SEM_INFO
    preco:       float | str = SEM_INFO
    descricao:   str = SEM_INFO
    link_obra:   str = SEM_INFO
    link_imagem: str = SEM_INFO


def parse_detalhe(html: str, url: str) -> Obra:
    soup = BeautifulSoup(html, "lxml")
    jsons = jsonld_from_soup(soup)
    texto_pagina = soup.get_text(" ", strip=True)

    obra = Obra(link_obra=url)

    # ── Título ──
    for seletor in [
        "h1.product_title",
        "h1.entry-title",
        ".product-title h1",
        'meta[property="og:title"]',
        "h1",
    ]:
        tag = soup.select_one(seletor)
        if tag:
            t = tag.get("content") if seletor.startswith("meta") else tag.get_text(" ", strip=True)
            if t and t.strip():
                obra.titulo = limpar(t)
                break
    if obra.titulo == SEM_INFO:
        t = jsonld_value(jsons, "name")
        if t != SEM_INFO:
            obra.titulo = t

    # ── Imagem ──
    for seletor in [
        ".woocommerce-product-gallery__image img",
        ".product-gallery img",
        ".product-image img",
        'meta[property="og:image"]',
        ".entry-content img",
        "figure img",
    ]:
        tag = soup.select_one(seletor)
        if tag:
            src = tag.get("content") if seletor.startswith("meta") else (
                tag.get("data-large_image") or tag.get("data-src") or tag.get("src")
            )
            if src and not src.startswith("data:"):
                obra.link_imagem = abs_url(src)
                break

    # ── Preço ──
    for seletor in [
        ".price ins .amount",
        ".price .woocommerce-Price-amount",
        ".price .amount",
        "p.price",
        ".product-price",
    ]:
        tag = soup.select_one(seletor)
        if tag:
            p = extrair_preco(tag.get_text(" ", strip=True))
            if p != SEM_INFO:
                obra.preco = p
                break
    if obra.preco == SEM_INFO:
        obra.preco = extrair_preco(texto_pagina)

    # ── Descrição ──
    for seletor in [
        ".woocommerce-product-details__short-description",
        ".product-description",
        ".entry-content",
        "[class*='description']",
        "div.product p",
    ]:
        tag = soup.select_one(seletor)
        if tag:
            t = tag.get_text(" ", strip=True)
            if t and len(t) > 10:
                obra.descricao = limpar(t[:1000])
                break

    # ── Metadados (ano, técnica, dimensões) ──
    # Tenta tabela de atributos (WooCommerce)
    for row in soup.select("table.shop_attributes tr, table.woocommerce-product-attributes tr"):
        label = limpar(row.select_one("th, td:first-child"))
        valor = limpar(row.select_one("td:last-child, td:nth-child(2)"))
        if label == SEM_INFO or valor == SEM_INFO:
            continue
        ll = label.lower()
        if any(k in ll for k in ("ano", "year", "data")):
            obra.ano = extrair_ano(valor) or valor
        elif any(k in ll for k in ("técnica", "tecnica", "medium", "material", "suporte")):
            obra.tecnica = valor
        elif any(k in ll for k in ("dimen", "medida", "tamanho", "size")):
            obra.dimensoes = padronizar_dimensoes(valor) or valor
        elif any(k in ll for k in ("preço", "preco", "price", "valor")):
            if obra.preco == SEM_INFO:
                obra.preco = extrair_preco(valor)

    # Fallback via texto
    if obra.ano == SEM_INFO:
        obra.ano = extrair_ano(texto_pagina)
    if obra.dimensoes == SEM_INFO:
        m = DIM_RE.search(texto_pagina)
        if m:
            obra.dimensoes = padronizar_dimensoes(m.group(0))

    # ── JSON-LD ──
    if obra.ano == SEM_INFO:
        v = jsonld_value(jsons, "dateCreated", "datePublished", "copyrightYear")
        if v != SEM_INFO:
            obra.ano = extrair_ano(v)
    if obra.descricao == SEM_INFO:
        v = jsonld_value(jsons, "description")
        if v != SEM_INFO:
            obra.descricao = limpar(v[:1000])

    return obra


# ─── Scraping principal ───────────────────────────────────────────────────────

def scrape_milan_dusek(
    url: str = ARTISTA_URL,
    max_obras: int = 0,
    delay: float = 0.4,
) -> pd.DataFrame:
    session = make_session()
    links = descobrir_links(session, url)
    if max_obras > 0:
        links = links[:max_obras]
    print(f"\nTotal de obras para detalhar: {len(links)}\n")

    rows: list[dict] = []
    for idx, link in enumerate(links, 1):
        print(f"[{idx}/{len(links)}] {link}")
        try:
            html = fetch(session, link)
            obra = parse_detalhe(html, link)
            rows.append({
                "Título":                 obra.titulo,
                "Autor":                  obra.autor,
                "Ano":                    obra.ano,
                "Técnica":                obra.tecnica,
                "Dimensões":              obra.dimensoes,
                "Preço":                  obra.preco,
                "Descrição":              obra.descricao,
                "Link da obra":           obra.link_obra,
                "Link da imagem da obra": obra.link_imagem,
            })
            print(f"  OK: {obra.titulo} | {obra.tecnica} | {obra.preco}")
        except Exception as e:
            print(f"  ERRO: {e}", file=sys.stderr)
            rows.append({
                "Título": SEM_INFO, "Autor": ARTISTA_NOME,
                "Ano": SEM_INFO, "Técnica": SEM_INFO, "Dimensões": SEM_INFO,
                "Preço": SEM_INFO, "Descrição": SEM_INFO,
                "Link da obra": link, "Link da imagem da obra": SEM_INFO,
            })
        if delay:
            time.sleep(delay)

    df = pd.DataFrame(rows, columns=EXCEL_COLUMNS)
    if not df.empty:
        df = df.drop_duplicates(subset=["Link da obra"], keep="first")
    return df


# ─── Exportação ───────────────────────────────────────────────────────────────

def exportar(df: pd.DataFrame, output: str) -> None:
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Dados")
        ws = writer.sheets["Dados"]
        for col in ws.columns:
            max_len = max((len(str(cell.value or "")) for cell in col), default=12)
            ws.column_dimensions[col[0].column_letter].width = min(max(max_len + 2, 12), 80)
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions


# ─── Entrypoint ───────────────────────────────────────────────────────────────

def main() -> None:
    parser = argparse.ArgumentParser(
        description=f"Scraper Arte Galeria — {ARTISTA_NOME}",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    parser.add_argument("--url",       default=os.getenv("ARTISTA_URL", ARTISTA_URL))
    parser.add_argument("--max-obras", type=int, default=int(os.getenv("MAX_OBRAS", "0")))
    parser.add_argument("--max-pages", type=int, default=0)
    parser.add_argument("--delay",     type=float, default=float(os.getenv("DELAY", "0.4")))
    parser.add_argument("--output",    default=os.getenv("OUTPUT", "artes_MilanDusek_ArteGaleria.xlsx"))
    args = parser.parse_args()

    iniciado = datetime.now()
    df = scrape_milan_dusek(url=args.url, max_obras=args.max_obras, delay=args.delay)
    exportar(df, args.output)
    elapsed = datetime.now() - iniciado
    print(f"\nConcluído: {len(df)} obras exportadas para '{args.output}' em {elapsed}.")


if __name__ == "__main__":
    main()
