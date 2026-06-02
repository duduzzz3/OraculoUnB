"""
Scraper Leilões BR — Pinturas e Gravuras
=========================================
Coleta obras da busca de Pinturas e Gravuras em andamento:
  https://leiloesbr.com.br/busca_andamento.asp?...tp=|50696E74757261732065204772617675726173|...

Estratégia principal
--------------------
Todo o dado rico (autor, título, técnica, dimensões, ano) já está no atributo
`alt` das imagens e no `title` dos links na página de listagem.  O scraper
extrai esses dados sem precisar abrir cada página de detalhe, tornando a coleta
muito mais rápida.  Quando --detalhe true é passado, abre cada obra
individualmente para tentar complementar dados faltantes.

Compatibilidade com o app Oráculo Cultural
------------------------------------------
    python leiloesbrWS_sem_miniaturas.py
    python leiloesbrWS_sem_miniaturas.py --max-obras 100 --output artes_LeiloesBR.xlsx --headless true
    python leiloesbrWS_sem_miniaturas.py --max-obras 0 --output saida.xlsx --detalhe false
    python leiloesbrWS_sem_miniaturas.py --html-file pagina_salva.html --output saida.xlsx
"""
from __future__ import annotations

import argparse
import os
import re
from pathlib import Path
import sys
import time
import unicodedata
from dataclasses import dataclass
from datetime import datetime
from typing import Any
from urllib.parse import urljoin, urlparse

import pandas as pd
import requests
from bs4 import BeautifulSoup, Tag
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

# ─── Constantes ───────────────────────────────────────────────────────────────
BASE_URL   = "https://leiloesbr.com.br"
SEARCH_URL = (
    f"{BASE_URL}/busca_andamento.asp"
    "?pesquisa=&op=2&v=126"
    "&tp=%7C50696E74757261732065204772617675726173%7C"
    "&b=0&pag={pag}"
)
SEM_INFO   = "Sem informação"

EXCEL_COLUMNS = [
    "Nome da obra", "Autor", "Preço", "Técnica", "Técnica original",
    "Dimensões", "Ano da obra", "Descrição",
    "Link da obra", "Link da imagem da obra",
]

# Padrões compilados
YEAR_RE  = re.compile(r"\b(1[5-9]\d{2}|20[0-2]\d)\b")
DIM_RE   = re.compile(
    r"\b(\d{1,4}(?:[.,]\d{1,2})?)\s*[xX×]\s*(\d{1,4}(?:[.,]\d{1,2})?)"
    r"(?:\s*[xX×]\s*(\d{1,4}(?:[.,]\d{1,2})?))?(?:\s*cm\b)?",
)
PRICE_RE = re.compile(
    r"R\$\s*(\d{1,3}(?:[.\s]\d{3})*(?:,\d{2})?|\d+(?:,\d{2})?)",
    re.I,
)
IMGS_INVALIDAS = re.compile(
    r"(logo|placeholder|banner|icon|icone|avatar|sem[\-_]imagem|"
    r"default|spacer|blank|loading|spinner|pixel\.gif|1x1)",
    re.I,
)
TITULOS_INVALIDOS = frozenset({
    "leilões", "conta", "login", "categorias", "0 lotes disponíveis",
    "pinturas e gravuras", "resultado da busca", "leilões br", "arremate",
    "sem título", "untitled", "leiloesbr", "todas as peças",
})

# Mapeamento de técnicas padronizadas
MAPA_TECNICA: list[tuple[str, str]] = [
    # Técnica mista (detectada antes de tudo)
    (r"t\.m(?:ista)?(?:\.|$)", "Técnica mista"),
    (r"t\.m\.s/(?:p|papel|tela)\b", "Técnica mista"),
    (r"mixed\s+media", "Técnica mista"),
    (r"técnica\s+mista", "Técnica mista"),
    # Óleos
    (r"o\.s/(?:tela|madeira|placa|papel|cart[aã]o|cort[aã]o|eucatex|eucatéx|hardboard|masonite|chapa|compensado|painel|tábua|cobre)", "Pintura a óleo"),
    (r"\b[oó]leo(?:\s+sobre|\s+s/)?", "Pintura a óleo"),
    (r"oil\s+on", "Pintura a óleo"),
    # Acrílico
    (r"a\.s/(?:tela|papel|madeira|placa|cart[aã]o)", "Pintura acrílica"),
    (r"acr[íi]l(?:ica?|ico?)\b", "Pintura acrílica"),
    (r"acrylic\b", "Pintura acrílica"),
    # Aquarela / Guache
    (r"\baquarela\b(?!da)", "Aquarela"),
    (r"waterh?colou?r\b", "Aquarela"),
    (r"\bguache\b", "Guache"),
    (r"\bgouache\b", "Guache"),
    # Nanquim
    (r"\bnanquim\b", "Nanquim / tinta"),
    (r"\bindia\s+ink\b", "Nanquim / tinta"),
    # Pastel
    (r"\bpastel\b", "Pastel seco"),
    # Carvão / Grafite
    (r"\bcarvão\b", "Carvão"),
    (r"\bcharcoal\b", "Carvão"),
    (r"\bgrafite\b", "Grafite"),
    (r"\bgraphite\b", "Grafite"),
    # Gravuras
    (r"\bserigrafia\b", "Serigrafia"),
    (r"\bserigraff?\b", "Serigrafia"),
    (r"\bscreenprint\b", "Serigrafia"),
    (r"\bsilkscreen\b", "Serigrafia"),
    (r"\bxilogravura\b", "Xilogravura"),
    (r"\bwoodcut\b", "Xilogravura"),
    (r"\blinoleogravura\b", "Linogravura"),
    (r"\blinoleum\b", "Linogravura"),
    (r"\blitogravura\b", "Litografia"),
    (r"\blitografia\b", "Litografia"),
    (r"\blithograph\b", "Litografia"),
    (r"\b[aá]gua[-\s]forte\b", "Água-forte"),
    (r"\betching\b", "Água-forte"),
    (r"\b[aá]gua[-\s]tinta\b", "Água-tinta"),
    (r"\baquatint\b", "Água-tinta"),
    (r"\bponta[-\s]seca\b", "Ponta-seca"),
    (r"\bdrypoint\b", "Ponta-seca"),
    (r"\bgravura(?:\s+em\s+metal)?\b", "Gravura em metal"),
    (r"\bengraving\b", "Gravura em metal"),
    (r"\bintaglio\b", "Gravura em metal"),
    # Desenho
    (r"\bdesenho\b", "Desenho"),
    (r"\bdrawing\b", "Desenho"),
    # Fotografia
    (r"\bfoto(?:grafia)?\b", "Fotografia"),
    (r"\bphotograph\b", "Fotografia"),
    # Escultura / volume
    (r"\bescultura\b", "Escultura"),
    (r"\bsculpture\b", "Escultura"),
    (r"\bbronze\b", "Escultura"),
    (r"\bcerâmica\b", "Cerâmica"),
    (r"\bporcelana\b", "Porcelana"),
    (r"\btapeçaria\b", "Tapeçaria"),
    (r"\btextil\b", "Têxtil"),
    (r"\btécnica\s+", "Técnica mista"),
]


# ─── Utilitários ──────────────────────────────────────────────────────────────

def limpar_texto(valor: Any) -> str:
    if valor is None:
        return SEM_INFO
    try:
        t = str(valor).encode("utf-8", "replace").decode("utf-8")
        t = t.replace("\xa0", " ").replace("\u200b", "").replace("\ufeff", "")
        t = re.sub(r"\s+", " ", t).strip()
    except Exception:
        t = str(valor).strip()
    return t if t else SEM_INFO


def decodificar_html(data: bytes) -> str:
    for enc in ("utf-8", "windows-1252", "iso-8859-1", "latin-1"):
        try:
            return data.decode(enc)
        except (UnicodeDecodeError, LookupError):
            continue
    return data.decode("utf-8", errors="replace")


def url_absoluta(href: str | None) -> str:
    if not href:
        return SEM_INFO
    href = href.strip()
    if href.startswith("javascript:") or href == "#":
        return SEM_INFO
    if href.startswith("//"):
        return "https:" + href
    if href.startswith("http"):
        return href
    return urljoin(BASE_URL, href)


def imagem_valida(url: str) -> bool:
    if not url or url == SEM_INFO:
        return False
    if IMGS_INVALIDAS.search(url):
        return False
    parsed = urlparse(url)
    if not parsed.scheme or not parsed.netloc:
        return False
    return bool(
        re.search(r"\.(jpg|jpeg|png|webp|gif)", parsed.path, re.I)
        or "cloudfront.net" in parsed.netloc
        or "cdn." in parsed.netloc
    )


def titulo_valido(titulo: Any) -> bool:
    t = limpar_texto(titulo)
    if t == SEM_INFO or len(t) < 3:
        return False
    tl = t.lower().strip()
    if tl in TITULOS_INVALIDOS:
        return False
    return not any(x in tl for x in ("leilões", "login", "cadastro", "busca", "r$", ".com"))


def normalizar_preco(texto: Any) -> float | str:
    m = PRICE_RE.search(str(texto or ""))
    if not m:
        return SEM_INFO
    raw = m.group(1).replace(".", "").replace(",", ".").replace(" ", "")
    try:
        return round(float(raw), 2)
    except ValueError:
        return SEM_INFO


def _fmt_dim(v: str) -> str:
    try:
        n = float(v.replace(",", "."))
        return str(int(n)) if n.is_integer() else str(round(n, 2)).rstrip("0").rstrip(".")
    except ValueError:
        return v.replace(",", ".").strip()


def normalizar_dimensoes(texto: str) -> str:
    t = str(texto or "")
    m = DIM_RE.search(t)
    if not m:
        return SEM_INFO
    nums = [float(x.replace(",", ".")) for x in [m.group(1), m.group(2)] if x]
    if m.group(3):
        nums.append(float(m.group(3).replace(",", ".")))
    if any(n > 2500 or n == 0 for n in nums):
        return SEM_INFO
    if len(nums) < 2:
        return SEM_INFO
    partes = [_fmt_dim(m.group(1)), _fmt_dim(m.group(2))]
    if m.group(3):
        partes.append(_fmt_dim(m.group(3)))
    return " x ".join(partes) + " cm"


def _expandir_ano_2digitos(texto: str) -> str:
    """Expande anos de 2 dígitos após 'datado/datada': datado 97 → datado 1997."""
    def _conv(m: re.Match) -> str:
        d = int(m.group(2))
        full = ("19" if d > 20 else "20") + m.group(2)
        return m.group(1) + full
    return re.sub(r"\b(datad[ao]\s+)(\d{2})\b", _conv, texto, flags=re.I)


def normalizar_ano(texto: str) -> str:
    """Extrai ano real da obra, removendo biografia e expandindo 2 dígitos."""
    t = str(texto or "")
    # Remove datas biográficas: (1935/2017), (1933..2025), (1935-2025), (1959)
    t = re.sub(r"\(\d{4}\s*[/\-\.]+\s*\d{4}\)", "", t)
    t = re.sub(r"\(\d{4}[/\-]\s*\.\.\)", "", t)
    t = re.sub(r"\(\d{4}[/\-]\s*\?\)", "", t)
    t = re.sub(r"\(\d{4}\.{2,}\d{4}\)", "", t)
    t = re.sub(r"\(\d{4}\)", "", t)          # (1959) isolado
    # Expande 2 dígitos
    t = _expandir_ano_2digitos(t)
    # Procura ano explicitamente mencionado
    m_exp = re.search(
        r"\b(?:datad[ao]|ano|year|from|de)\s+(1[5-9]\d{2}|20[0-2]\d)\b",
        t, re.I,
    )
    if m_exp:
        return m_exp.group(1)
    candidatos = YEAR_RE.findall(t)
    if not candidatos:
        return SEM_INFO
    ano_atual = datetime.now().year
    validos = [a for a in candidatos if 1700 <= int(a) <= ano_atual + 1]
    if not validos:
        return SEM_INFO
    return min(validos, key=int)


def normalizar_tecnica(texto_original: str) -> str:
    """Mapeia texto bruto de técnica para categoria padronizada."""
    t = limpar_texto(texto_original)
    if t == SEM_INFO:
        return SEM_INFO
    tl = t.lower()
    for padrao, padronizada in MAPA_TECNICA:
        try:
            if re.search(padrao, tl, re.I):
                return padronizada
        except re.error:
            pass
    return "Outros"


# ─── Parsing da descrição do card ─────────────────────────────────────────────

# Abreviações a buscar na descrição (em lowercase)
_ABREVS_ORDENADAS: list[str] = [
    "o.s/tela", "o.s/madeira", "o.s/placa", "o.s/papel", "o.s/cart",
    "o.s/euca", "o.s/hards", "o.s/chapa", "o.s/comp", "o.s/pan",
    "o.s/táb", "o.s/tab", "o.s/cob",
    "a.s/tela", "a.s/papel", "a.s/madeira", "a.s/",
    "t.m.s/p", "t.m.s/t", "t.m.", "t.mista",
    "serigrafia", "xilogravura", "litografia", "litogravura",
    "linoleogravura", "nanquim", "aquarela", "guache",
    "água-forte", "agua-forte", "água-tinta", "agua-tinta",
    "ponta-seca", "etching", "aquatint",
    "pastel", "carvão", "grafite", "desenho", "fotografia",
    "escultura", "bronze", "cerâmica", "porcelana", "tapeçaria",
    "gravura", "monotipia",
]

# Seletores que delimitam FIM do título (começa o bloco de metadados)
_RE_INICIO_META = re.compile(
    r"\b(?:o\.s/|a\.s/|t\.m(?:\.|ista)|serigrafia|xilogravura|litografia|litogravura|"
    r"linoleogravura|nanquim|aquarela|guache|[aá]gua[-]?(?:forte|tinta)|ponta[-]?seca|"
    r"gravura|monotipia|pastel|carvão|grafite|desenho|fotografia|escultura|bronze|"
    r"cerâmica|porcelana|tapeçaria)\b",
    re.I,
)


def _extrair_abrev_tecnica(texto: str) -> str:
    """Retorna o trecho de técnica encontrado na descrição (partindo do início da abrev)."""
    tl = texto.lower()
    for ab in _ABREVS_ORDENADAS:
        idx = tl.find(ab)
        if idx >= 0:
            # Captura a partir do índice da abreviação (não antes, para evitar vírgula inicial)
            trecho = texto[idx:idx + len(ab) + 35].strip()
            # Para na próxima vírgula, ponto-e-vírgula ou parêntese
            trecho = re.split(r"[,;(]", trecho)[0].strip().rstrip(".")
            return trecho
    return SEM_INFO


def _parse_descricao_card(desc: str) -> dict:
    """
    Analisa a descrição rica do alt/title dos cards do Leilões BR.

    Formato típico:
        AUTOR SOBRENOME (ano_nasc/ano_morte),  Título da obra, técnica, AxBcm, info...

    Retorna: autor, titulo, tecnica_original, dimensoes, ano
    """
    desc = limpar_texto(desc)
    if desc == SEM_INFO:
        return {}

    resultado = {
        "autor":            SEM_INFO,
        "titulo":           SEM_INFO,
        "tecnica_original": SEM_INFO,
        "dimensoes":        SEM_INFO,
        "ano":              SEM_INFO,
    }

    # ── Dimensões ──────────────────────────────────────────────────────────────
    dim = normalizar_dimensoes(desc)
    if dim != SEM_INFO:
        resultado["dimensoes"] = dim

    # ── Técnica (bruta) ─────────────────────────────────────────────────────
    tec_bruta = _extrair_abrev_tecnica(desc)
    if tec_bruta != SEM_INFO:
        resultado["tecnica_original"] = tec_bruta

    # ── Ano (removendo bio) ────────────────────────────────────────────────
    ano = normalizar_ano(desc)
    if ano != SEM_INFO:
        resultado["ano"] = ano

    # ── Autor e Título ─────────────────────────────────────────────────────
    # Separa no primeiro ", " que divide AUTOR do resto
    m_autor = re.match(
        r"^([A-ZÀÁÂÃÄÅÈÉÊËÌÍÎÏÒÓÔÕÖÙÚÛÜÇÑ][^,]{1,79}?)\s*,\s+(.+)$",
        desc, re.S,
    )
    if m_autor:
        candidato_autor = m_autor.group(1).strip()
        resto = m_autor.group(2).strip()

        # Valida: tem letras maiúsculas e não é número/dimensão
        if (
            re.search(r"[A-ZÁÉÍÓÚÀÂÊÎÔÛÃÕÇ]{2,}", candidato_autor)
            and not re.search(r"^\d", candidato_autor)
            and len(candidato_autor) <= 80
        ):
            # Remove datas biográficas do nome
            autor_limpo = re.sub(r"\s*\(\s*\d{4}\s*[./\-]{1,2}\s*[\d.?]+\s*\)", "", candidato_autor).strip()
            autor_limpo = re.sub(r"\s*\(\s*\?\s*\)", "", autor_limpo).strip()
            autor_limpo = re.sub(r"\s*\(\d{4}\)", "", autor_limpo).strip()
            resultado["autor"] = autor_limpo.title()

            # Título: texto antes da técnica ou da dimensão
            m_tec_inicio = _RE_INICIO_META.search(resto)
            if m_tec_inicio:
                titulo_bruto = resto[:m_tec_inicio.start()].strip().rstrip(",").strip()
            else:
                m_dim2 = DIM_RE.search(resto)
                titulo_bruto = (
                    resto[:m_dim2.start()].strip().rstrip(",").strip()
                    if m_dim2 else re.split(r",", resto)[0].strip()
                )

            if titulo_valido(titulo_bruto):
                resultado["titulo"] = titulo_bruto
    else:
        if titulo_valido(desc):
            resultado["titulo"] = desc

    return resultado


# ─── Sessão HTTP ──────────────────────────────────────────────────────────────

def criar_sessao() -> requests.Session:
    retry = Retry(
        total=5, connect=4, read=4, status=3,
        backoff_factor=1.2,
        status_forcelist=(429, 500, 502, 503, 504),
        allowed_methods=("GET",),
        raise_on_status=False,
    )
    s = requests.Session()
    s.mount("https://", HTTPAdapter(max_retries=retry, pool_connections=4, pool_maxsize=4))
    s.mount("http://",  HTTPAdapter(max_retries=retry, pool_connections=4, pool_maxsize=4))
    s.headers.update({
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/124.0.0.0 Safari/537.36"
        ),
        "Accept":          "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8",
        "Accept-Language": "pt-BR,pt;q=0.9,en-US;q=0.8,en;q=0.7",
        "Accept-Encoding": "gzip, deflate, br",
        "Connection":      "keep-alive",
        "Upgrade-Insecure-Requests": "1",
        "Sec-Fetch-Dest":  "document",
        "Sec-Fetch-Mode":  "navigate",
        "Sec-Fetch-Site":  "same-origin",
        "Sec-Fetch-User":  "?1",
        "sec-ch-ua":       '"Chromium";v="124", "Google Chrome";v="124", "Not-A.Brand";v="99"',
        "sec-ch-ua-mobile":   "?0",
        "sec-ch-ua-platform": '"Windows"',
    })
    return s


def _buscar_com_requests(session: requests.Session, url: str, timeout: int = 30) -> BeautifulSoup | None:
    """Tenta buscar via requests com warm-up no homepage para obter cookies."""
    try:
        # Warm-up para obter cookies de sessão
        session.headers["Referer"] = BASE_URL + "/"
        resp = session.get(url, timeout=timeout, allow_redirects=True)
        if resp.status_code == 403:
            return None
        resp.raise_for_status()
        html = decodificar_html(resp.content)
        if "abre_catalogo" not in html and "Itens encontrados" not in html:
            return None
        return BeautifulSoup(html, "lxml")
    except Exception:
        return None


def _buscar_com_playwright(url: str, headless: bool = True, timeout_ms: int = 30_000) -> BeautifulSoup | None:
    """Fallback com Playwright para sites que bloqueiam requests."""
    try:
        from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout
    except ImportError:
        print("  [AVISO] Playwright não instalado. Instale com: pip install playwright && playwright install chromium", file=sys.stderr)
        return None
    try:
        with sync_playwright() as pw:
            browser = pw.chromium.launch(
                headless=headless,
                args=["--disable-blink-features=AutomationControlled", "--no-sandbox",
                      "--disable-dev-shm-usage"],
            )
            ctx = browser.new_context(
                viewport={"width": 1440, "height": 900},
                locale="pt-BR",
                user_agent=(
                    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"
                ),
            )
            page = ctx.new_page()
            # Warm-up homepage para cookies
            try:
                page.goto(BASE_URL + "/default.asp", wait_until="domcontentloaded", timeout=timeout_ms)
                page.wait_for_timeout(800)
            except Exception:
                pass
            page.goto(url, wait_until="domcontentloaded", timeout=timeout_ms)
            page.wait_for_timeout(1_200)
            html = page.content()
            ctx.close()
            browser.close()
        return BeautifulSoup(html, "lxml")
    except Exception as e:
        print(f"  [PLAYWRIGHT] Erro: {e}", file=sys.stderr)
        return None


def buscar_pagina(
    session: requests.Session,
    url: str,
    usar_playwright: bool = False,
    headless: bool = True,
) -> BeautifulSoup | None:
    """Busca com requests; faz fallback para Playwright se retornar 403."""
    if not usar_playwright:
        soup = _buscar_com_requests(session, url)
        if soup is not None:
            return soup
        print(f"  [INFO] requests bloqueado, tentando Playwright para: {url}")

    return _buscar_com_playwright(url, headless=headless)


# ─── Extração dos links de obras ──────────────────────────────────────────────

def extrair_links_obras(soup: BeautifulSoup) -> list[tuple[str, str, str]]:
    """
    Retorna lista de (url_obra, descricao, url_imagem).

    Padrão Leilões BR:
        <a href="abre_catalogo.asp?t=1|URL_GALERIA|ID|ID" title="DESC COMPLETA">
          <img alt="DESC COMPLETA" src="URL_CDN">
        </a>
    """
    resultados: list[tuple[str, str, str]] = []
    vistos: set[str] = set()

    for tag_a in soup.select("a[href*='abre_catalogo.asp']"):
        href = tag_a.get("href", "").strip()
        if not href or href in vistos:
            continue
        link = url_absoluta(href)
        if link == SEM_INFO:
            continue

        desc = limpar_texto(tag_a.get("title", ""))
        img_url = SEM_INFO
        img_tag = tag_a.find("img")
        if img_tag:
            if desc == SEM_INFO:
                desc = limpar_texto(img_tag.get("alt", ""))
            for attr in ("src", "data-src", "data-original", "data-lazy"):
                src = img_tag.get(attr, "")
                if src and not src.startswith("data:"):
                    u = url_absoluta(src.strip())
                    if imagem_valida(u):
                        img_url = u
                        break

        # Filtra links institucionais
        if not desc or desc == SEM_INFO:
            continue
        if any(x in desc.lower() for x in ("leilões br", "logo", "banner")):
            continue

        vistos.add(href)
        resultados.append((link, desc, img_url))

    return resultados


def extrair_preco_card(tag_a: Tag) -> float | str:
    """Extrai preço do container do card."""
    container = tag_a.find_parent(["li", "div", "article", "td"]) or tag_a
    return normalizar_preco(container.get_text(" ", strip=True))


# ─── Extração de detalhe (opcional) ──────────────────────────────────────────

def extrair_dados_obra_detalhe(
    session: requests.Session,
    url: str,
    headless: bool = True,
) -> dict:
    """Visita página individual para complementar dados."""
    dados: dict = {}
    soup = buscar_pagina(session, url, headless=headless)
    if not soup:
        return dados

    texto = soup.get_text(" ", strip=True)

    og_img = soup.find("meta", property="og:image")
    if og_img:
        u = url_absoluta(limpar_texto(og_img.get("content", "")))
        if imagem_valida(u):
            dados["link_imagem"] = u

    for sel in ["meta[property='og:title']", "h1.titulo", "h1.lote-titulo", "h1"]:
        t_tag = soup.select_one(sel)
        if t_tag:
            t = limpar_texto(t_tag.get("content") if sel.startswith("meta") else t_tag.get_text())
            if titulo_valido(t):
                dados["titulo"] = t
                break

    for sel in [".preco", ".lance", ".valor", "[class*='price']", "[class*='lance']"]:
        t_tag = soup.select_one(sel)
        if t_tag:
            p = normalizar_preco(t_tag.get_text())
            if p != SEM_INFO:
                dados["preco"] = p
                break
    if "preco" not in dados:
        p = normalizar_preco(texto)
        if p != SEM_INFO:
            dados["preco"] = p

    return dados



# ─── Modo anti-403 legítimo: HTML salvo localmente ───────────────────────────

def soup_de_arquivo_html(path: str | Path) -> BeautifulSoup | None:
    """Lê HTML salvo pelo navegador e devolve BeautifulSoup.

    Este modo é útil quando o Leilões BR retorna 403 para o Streamlit Cloud
    ou para requests, mas a página abre normalmente no navegador do usuário.
    Salve a página como HTML e execute com --html-file ou --html-dir.
    """
    p = Path(path)
    if not p.exists() or not p.is_file():
        print(f"  [HTML] Arquivo não encontrado: {p}", file=sys.stderr)
        return None
    data = p.read_bytes()
    html = decodificar_html(data)
    if "abre_catalogo" not in html and "Itens encontrados" not in html and "peca.asp" not in html:
        print(f"  [HTML] Aviso: {p.name} não parece ser uma página de listagem/lote do Leilões BR.")
    return BeautifulSoup(html, "lxml")


def listar_arquivos_html(html_file: str | None = None, html_dir: str | None = None) -> list[Path]:
    arquivos: list[Path] = []
    if html_file:
        arquivos.append(Path(html_file))
    if html_dir:
        d = Path(html_dir)
        if d.exists() and d.is_dir():
            arquivos.extend(sorted([*d.glob("*.html"), *d.glob("*.htm")]))
    # Remove duplicatas preservando ordem
    vistos: set[str] = set()
    out: list[Path] = []
    for p in arquivos:
        key = str(p.resolve()) if p.exists() else str(p)
        if key not in vistos:
            vistos.add(key)
            out.append(p)
    return out


def _obras_de_soup_listagem(
    soup: BeautifulSoup,
    links_vistos: set[str],
    max_restante: int | None = None,
) -> list[Obra]:
    """Extrai obras de uma página de listagem já carregada."""
    obras: list[Obra] = []
    pares = extrair_links_obras(soup)
    novos = [(l, d, i) for l, d, i in pares if l not in links_vistos]
    print(f"  [PARSE] {len(novos)} links novos ({len(pares)} encontrados no HTML)")

    tags_por_link: dict[str, Tag] = {}
    for tag_a in soup.select("a[href*='abre_catalogo.asp']"):
        href = tag_a.get("href", "")
        l_abs = url_absoluta(href)
        if l_abs not in tags_por_link:
            tags_por_link[l_abs] = tag_a

    for link, desc, img_url_card in novos:
        if max_restante is not None and len(obras) >= max_restante:
            break
        links_vistos.add(link)
        obra = Obra(link_obra=link, descricao=limpar_texto(desc))

        d = _parse_descricao_card(desc)
        if d.get("autor") and d["autor"] != SEM_INFO:
            obra.autor = d["autor"]
        if d.get("titulo") and titulo_valido(d["titulo"]):
            obra.nome_obra = d["titulo"]
        if d.get("dimensoes") and d["dimensoes"] != SEM_INFO:
            obra.dimensoes = d["dimensoes"]
        if d.get("ano") and d["ano"] != SEM_INFO:
            obra.ano_obra = d["ano"]
        if d.get("tecnica_original") and d["tecnica_original"] != SEM_INFO:
            obra.tecnica_original = d["tecnica_original"]
            obra.tecnica = normalizar_tecnica(obra.tecnica_original)
        if imagem_valida(img_url_card):
            obra.link_imagem = img_url_card

        tag_a = tags_por_link.get(link)
        if tag_a:
            p = extrair_preco_card(tag_a)
            if p != SEM_INFO:
                obra.preco = p

        obras.append(obra)
        print(
            f"  [{len(obras):>4}] {obra.nome_obra[:45]!r:47s} "
            f"| {obra.autor[:22]!r:24s} "
            f"| {str(obra.preco):>10} "
            f"| {obra.tecnica}"
        )
    return obras


def scrape_leiloesbr_html(
    html_files: list[str | Path],
    max_obras: int = 0,
) -> pd.DataFrame:
    """Extrai obras de páginas HTML salvas localmente.

    Use quando o servidor bloquear a coleta automática com 403. O conteúdo é
    analisado pelo mesmo parser do scraper ao vivo.
    """
    obras: list[Obra] = []
    links_vistos: set[str] = set()
    for p in html_files:
        if max_obras > 0 and len(obras) >= max_obras:
            break
        print(f"[HTML] Lendo: {p}")
        soup = soup_de_arquivo_html(p)
        if not soup:
            continue
        restante = None if max_obras <= 0 else max_obras - len(obras)
        obras.extend(_obras_de_soup_listagem(soup, links_vistos, max_restante=restante))
    print(f"\nTotal HTML: {len(obras)} obras ({len(links_vistos)} links únicos).")
    return _obras_para_dataframe(obras)


# ─── Scraping principal ───────────────────────────────────────────────────────

@dataclass
class Obra:
    nome_obra:        str = "Sem título"
    autor:            str = SEM_INFO
    preco:            Any = SEM_INFO
    tecnica:          str = SEM_INFO
    tecnica_original: str = SEM_INFO
    dimensoes:        str = SEM_INFO
    ano_obra:         str = SEM_INFO
    descricao:        str = SEM_INFO
    link_obra:        str = SEM_INFO
    link_imagem:      str = SEM_INFO


def scrape_leiloesbr(
    max_obras: int = 0,
    visitar_detalhe: bool = False,
    delay: float = 0.3,
    headless: bool = True,
    start_page: int = 1,
    max_pages: int = 0,
) -> pd.DataFrame:
    session  = criar_sessao()
    obras: list[Obra] = []
    links_vistos: set[str] = set()
    pagina = max(1, int(start_page))
    paginas_processadas = 0
    usar_playwright = False  # ligado automaticamente após primeiro 403

    # Warm-up: busca homepage para obter cookies
    try:
        r0 = session.get(BASE_URL + "/default.asp", timeout=15)
        session.headers["Referer"] = BASE_URL + "/default.asp"
    except Exception:
        pass

    while True:
        if max_pages > 0 and paginas_processadas >= max_pages:
            print(f"Limite de {max_pages} página(s) atingido.")
            break
        paginas_processadas += 1
        url_pag = SEARCH_URL.format(pag=pagina)
        print(f"Página {pagina}: buscando {url_pag}")
        soup = buscar_pagina(session, url_pag, usar_playwright=usar_playwright, headless=headless)

        if soup is None:
            if not usar_playwright:
                print("  [INFO] Ativando Playwright para contornar bloqueio...")
                usar_playwright = True
                soup = buscar_pagina(session, url_pag, usar_playwright=True, headless=headless)
            if soup is None:
                print(f"  [ERRO] Falha na página {pagina}. Encerrando.")
                break

        # Extrai links + imagens + descrições
        pares = extrair_links_obras(soup)
        novos = [(l, d, i) for l, d, i in pares if l not in links_vistos]
        print(f"Página {pagina}: {len(novos)} links novos ({len(pares)} encontrados no total)")

        if not novos:
            print("Sem novos links. Fim da paginação.")
            break

        # Mapeia link → tag_a para extrair preço do card
        tags_por_link: dict[str, Tag] = {}
        for tag_a in soup.select("a[href*='abre_catalogo.asp']"):
            href = tag_a.get("href", "")
            l_abs = url_absoluta(href)
            if l_abs not in tags_por_link:
                tags_por_link[l_abs] = tag_a

        for link, desc, img_url_card in novos:
            if max_obras > 0 and len(obras) >= max_obras:
                print(f"Limite de {max_obras} obras atingido.")
                break

            links_vistos.add(link)
            obra = Obra(link_obra=link, descricao=limpar_texto(desc))

            # Dados do card
            d = _parse_descricao_card(desc)
            if d.get("autor") and d["autor"] != SEM_INFO:
                obra.autor = d["autor"]
            if d.get("titulo") and titulo_valido(d["titulo"]):
                obra.nome_obra = d["titulo"]
            if d.get("dimensoes") and d["dimensoes"] != SEM_INFO:
                obra.dimensoes = d["dimensoes"]
            if d.get("ano") and d["ano"] != SEM_INFO:
                obra.ano_obra = d["ano"]
            if d.get("tecnica_original") and d["tecnica_original"] != SEM_INFO:
                obra.tecnica_original = d["tecnica_original"]
                obra.tecnica = normalizar_tecnica(obra.tecnica_original)

            # Imagem do card
            if imagem_valida(img_url_card):
                obra.link_imagem = img_url_card

            # Preço do card
            tag_a = tags_por_link.get(link)
            if tag_a:
                p = extrair_preco_card(tag_a)
                if p != SEM_INFO:
                    obra.preco = p

            # Detalhe opcional
            if visitar_detalhe:
                try:
                    det = extrair_dados_obra_detalhe(session, link, headless=headless)
                    if det.get("link_imagem") and obra.link_imagem == SEM_INFO:
                        obra.link_imagem = det["link_imagem"]
                    if det.get("titulo") and not titulo_valido(obra.nome_obra):
                        obra.nome_obra = det["titulo"]
                    if det.get("preco") and obra.preco == SEM_INFO:
                        obra.preco = det["preco"]
                    time.sleep(delay)
                except Exception as e:
                    print(f"  [AVISO] Detalhe falhou {link}: {e}", file=sys.stderr)

            obras.append(obra)
            print(
                f"  [{len(obras):>4}] {obra.nome_obra[:45]!r:47s} "
                f"| {obra.autor[:22]!r:24s} "
                f"| {str(obra.preco):>10} "
                f"| {obra.tecnica}"
            )

        if max_obras > 0 and len(obras) >= max_obras:
            break

        # Detecta próxima página
        prox = _detectar_proxima_pagina(soup, pagina)
        if prox is None:
            print("Fim da paginação.")
            break
        pagina = prox
        time.sleep(delay)

    print(f"\nTotal: {len(obras)} obras ({len(links_vistos)} links únicos visitados).")
    return _obras_para_dataframe(obras)


def _detectar_proxima_pagina(soup: BeautifulSoup, pagina_atual: int) -> int | None:
    # Links pag=N
    for a in soup.select("a[href*='pag=']"):
        m = re.search(r"pag=(\d+)", a.get("href", ""))
        if m and int(m.group(1)) == pagina_atual + 1:
            return pagina_atual + 1
    # Links "próximo" / ">"
    for a in soup.find_all("a"):
        txt = a.get_text(" ", strip=True).lower()
        if txt in (">", "próximo", "próxima", "next", "»", "›"):
            m = re.search(r"pag=(\d+)", a.get("href", ""))
            if m:
                return int(m.group(1))
    # Verifica se a próxima existe
    paginas = {int(m.group(1)) for a in soup.select("a[href*='pag=']")
               for m in [re.search(r"pag=(\d+)", a.get("href", ""))] if m}
    return pagina_atual + 1 if pagina_atual + 1 in paginas else None


def _obras_para_dataframe(obras: list[Obra]) -> pd.DataFrame:
    rows = [{
        "Nome da obra":           o.nome_obra,
        "Autor":                  o.autor,
        "Preço":                  o.preco,
        "Técnica":                o.tecnica,
        "Técnica original":       o.tecnica_original,
        "Dimensões":              o.dimensoes,
        "Ano da obra":            o.ano_obra,
        "Descrição":              o.descricao,
        "Link da obra":           o.link_obra,
        "Link da imagem da obra": o.link_imagem,
    } for o in obras]
    df = pd.DataFrame(rows, columns=EXCEL_COLUMNS)
    if not df.empty:
        df = df.drop_duplicates(subset=["Link da obra"], keep="first")
    return df


# ─── Exportação ───────────────────────────────────────────────────────────────

def salvar_excel(df: pd.DataFrame, output: str) -> None:
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Dados")
        ws = writer.sheets["Dados"]
        for col in ws.columns:
            w = max((len(str(c.value or "")) for c in col if c.value), default=12)
            ws.column_dimensions[col[0].column_letter].width = min(max(w + 2, 12), 80)
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
    print(f"Arquivo gerado: {output}")
    print(f"Obras salvas: {len(df)}")


# ─── Entrypoint ───────────────────────────────────────────────────────────────

def main() -> None:
    parser = argparse.ArgumentParser(
        description="Scraper Leilões BR — Pinturas e Gravuras",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    parser.add_argument("--max-obras", type=int,
                        default=int(os.environ.get("MAX_OBRAS", "0")),
                        help="Limite de obras (0 = sem limite)")
    parser.add_argument("--output",
                        default=os.environ.get("OUTPUT", "artes_LeiloesBR.xlsx"),
                        help="Arquivo Excel de saída")
    parser.add_argument("--headless", default="true", choices=["true", "false"],
                        help="Rodar navegador invisível (Playwright fallback)")
    parser.add_argument("--detalhe", default="false", choices=["true", "false"],
                        help="Visitar página individual de cada obra")
    parser.add_argument("--delay", type=float,
                        default=float(os.environ.get("DELAY", "0.3")),
                        help="Pausa em segundos entre requisições")
    parser.add_argument("--start-page", type=int, default=int(os.environ.get("START_PAGE", "1")),
                        help="Página inicial da busca")
    parser.add_argument("--max-pages", type=int, default=int(os.environ.get("MAX_PAGES", "0")),
                        help="Máximo de páginas da busca (0 = sem limite)")
    parser.add_argument("--html-file", default=os.environ.get("HTML_FILE"),
                        help="Arquivo HTML salvo pelo navegador para extrair offline")
    parser.add_argument("--html-dir", default=os.environ.get("HTML_DIR"),
                        help="Pasta com arquivos .html/.htm salvos pelo navegador")
    args = parser.parse_args()

    inicio = datetime.now()
    html_files = listar_arquivos_html(args.html_file, args.html_dir)
    if html_files:
        print("Modo HTML offline ativado. Nenhuma requisição será feita ao Leilões BR.")
        df = scrape_leiloesbr_html(html_files, max_obras=args.max_obras)
    else:
        df = scrape_leiloesbr(
            max_obras=args.max_obras,
            visitar_detalhe=args.detalhe.lower() == "true",
            delay=args.delay,
            headless=args.headless.lower() == "true",
            start_page=args.start_page,
            max_pages=args.max_pages,
        )
    salvar_excel(df, args.output)
    print(f"Tempo total: {datetime.now() - inicio}")


if __name__ == "__main__":
    main()
