"""
Scraper Arremate Arte – Rubem Grilo
Coleta obras/lotes de https://www.arrematearte.com.br/artistas/rubem-grilo-1946

Usa Playwright (Chromium) para lidar com carregamento dinâmico do site de leilão.

Uso básico:
    python rubem_grilo_arremateWS.py
    python rubem_grilo_arremateWS.py --max-obras 60 --output saida.xlsx --headless true
"""
from __future__ import annotations

import argparse
import json
import os
import re
import sys
import time
from dataclasses import dataclass
from datetime import datetime
from typing import Any
from urllib.parse import urljoin

import pandas as pd
from playwright.sync_api import (
    Browser, BrowserContext, Page,
    TimeoutError as PlaywrightTimeoutError,
    sync_playwright,
)

# ─── Constantes ───────────────────────────────────────────────────────────────
BASE_URL    = "https://www.arrematearte.com.br"
ARTISTA_URL = f"{BASE_URL}/artistas/rubem-grilo-1946"
ARTISTA_NOME = "Rubem Grilo"
SEM_INFO    = "Sem informação"

EXCEL_COLUMNS = [
    "Título", "Autor", "Ano", "Técnica", "Dimensões",
    "Preço", "Descrição", "Link da obra", "Link da imagem da obra",
]

YEAR_RE  = re.compile(r"\b(19|20)\d{2}\b")
DIM_RE   = re.compile(
    r"\b(\d{1,4}(?:[.,]\d{1,2})?)\s*(?:x|×|X|por)\s*(\d{1,4}(?:[.,]\d{1,2})?)(?:\s*(?:x|×|X|por)\s*(\d{1,4}(?:[.,]\d{1,2})?))?(?:\s*cm\b)?",
    re.I,
)
PRICE_RE = re.compile(
    r"(?:R\$|BRL)\s*(\d{1,3}(?:[.\s]\d{3})*(?:,\d{2})?|\d+(?:,\d{2})?)"
    r"|(\d{1,3}(?:[.\s]\d{3})*(?:,\d{2})?)\s*(?:reais?)",
    re.I,
)


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
    if href.startswith("/"):
        return BASE_URL + href
    if href.startswith("http"):
        return href
    return SEM_INFO


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
    raw = (m.group(1) or m.group(2) or "").replace(".", "").replace(",", ".").replace(" ", "")
    try:
        return round(float(raw), 2)
    except ValueError:
        return SEM_INFO


def texto_loc(page_or_locator, seletor: str) -> str:
    """Extrai texto de um seletor de forma segura."""
    try:
        loc = page_or_locator.locator(seletor).first
        if loc.count() > 0:
            t = loc.inner_text().strip()
            return t if t else SEM_INFO
    except Exception:
        pass
    return SEM_INFO


def attr_loc(page_or_locator, seletor: str, attr: str) -> str:
    """Extrai atributo de um seletor de forma segura."""
    try:
        loc = page_or_locator.locator(seletor).first
        if loc.count() > 0:
            v = loc.get_attribute(attr)
            return v if v else SEM_INFO
    except Exception:
        pass
    return SEM_INFO


# ─── Extração de detalhes ─────────────────────────────────────────────────────

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


def extrair_imagem_detalhe(page: Page) -> str:
    """Tenta obter a melhor URL de imagem disponível na página de detalhe."""
    for seletor in [
        'meta[property="og:image"]',
        ".lot-image img",
        ".product-image img",
        ".artwork-image img",
        ".obra-imagem img",
        ".gallery-image img",
        "figure img",
        ".carousel img",
        "img[class*='main']",
        "img[class*='principal']",
    ]:
        try:
            loc = page.locator(seletor).first
            if loc.count() > 0:
                src = (
                    loc.get_attribute("content")
                    or loc.get_attribute("data-src")
                    or loc.get_attribute("data-original")
                    or loc.get_attribute("src")
                )
                if src and not src.startswith("data:"):
                    return abs_url(src)
        except Exception:
            pass
    return SEM_INFO


def _parse_metadados_texto(texto: str, obra: Obra) -> None:
    """Extrai metadados (ano, técnica, dimensões, preço) a partir do texto livre."""
    linhas = [limpar(l) for l in texto.splitlines() if l.strip()]
    for idx, linha in enumerate(linhas):
        ll = linha.lower()
        prox = linhas[idx + 1] if idx + 1 < len(linhas) else ""

        if obra.ano == SEM_INFO:
            if any(k in ll for k in ("ano:", "year:", "data:", "período:")):
                obra.ano = extrair_ano(prox or linha)
            elif re.search(r"^\s*(?:ano|year)\s*:?\s*", ll):
                obra.ano = extrair_ano(linha)

        if obra.tecnica == SEM_INFO:
            if any(k in ll for k in ("técnica:", "tecnica:", "medium:", "material:", "suporte:")):
                obra.tecnica = limpar(prox or linha.split(":", 1)[-1])

        if obra.dimensoes == SEM_INFO:
            if any(k in ll for k in ("dimensões:", "dimensoes:", "medidas:", "tamanho:", "size:")):
                candidato = prox or linha.split(":", 1)[-1]
                d = padronizar_dimensoes(candidato)
                if d != SEM_INFO:
                    obra.dimensoes = d
            # Tenta encontrar dimensões em qualquer linha
            if obra.dimensoes == SEM_INFO:
                d = padronizar_dimensoes(linha)
                if d != SEM_INFO:
                    obra.dimensoes = d

        if obra.preco == SEM_INFO:
            if any(k in ll for k in ("preço", "valor", "lance", "arrematado", "estimativa")):
                p = extrair_preco(prox or linha)
                if p != SEM_INFO:
                    obra.preco = p

    # Fallbacks globais
    if obra.ano == SEM_INFO:
        obra.ano = extrair_ano(texto)
    if obra.preco == SEM_INFO:
        obra.preco = extrair_preco(texto)


def extrair_detalhe(detail_page: Page, link: str) -> Obra:
    obra = Obra(link_obra=link)

    try:
        detail_page.goto(link, wait_until="domcontentloaded", timeout=40_000)
        detail_page.wait_for_timeout(1_500)

        # ── Título ──
        for seletor in [
            "h1.lot-title",
            "h1.artwork-title",
            "h1.obra-titulo",
            "h1.product-title",
            "h1.entry-title",
            "h1",
            'meta[property="og:title"]',
        ]:
            v = SEM_INFO
            if seletor.startswith("meta"):
                v = attr_loc(detail_page, seletor, "content")
            else:
                v = texto_loc(detail_page, seletor)
            if v != SEM_INFO and len(v) > 2:
                obra.titulo = v
                break

        # ── Imagem ──
        obra.link_imagem = extrair_imagem_detalhe(detail_page)

        # ── Preço ──
        for seletor in [
            ".lot-price",
            ".artwork-price",
            ".preco",
            ".price",
            ".valor-lance",
            ".valor-arrematado",
            ".estimativa",
            "[class*='price']",
            "[class*='preco']",
            "[class*='valor']",
        ]:
            v = texto_loc(detail_page, seletor)
            if v != SEM_INFO:
                p = extrair_preco(v)
                if p != SEM_INFO:
                    obra.preco = p
                    break

        # ── Metadados estruturados (tabelas/listas de detalhes) ──
        for seletor_label in [
            ".lot-details dt, .lot-details dd",
            ".artwork-details dt, .artwork-details dd",
            ".details-list li",
            "table.lot-info tr",
            ".product-meta tr",
            "[class*='detail'] [class*='label'], [class*='detail'] [class*='value']",
            ".field-label, .field-value",
        ]:
            try:
                items = detail_page.locator(seletor_label).all()
                if not items:
                    continue
                # Pares label/valor
                for j in range(0, len(items) - 1, 2):
                    label = items[j].inner_text().strip().lower()
                    valor = items[j + 1].inner_text().strip() if j + 1 < len(items) else ""
                    if not label or not valor:
                        continue
                    if any(k in label for k in ("ano", "year", "data")):
                        obra.ano = extrair_ano(valor) or obra.ano
                    elif any(k in label for k in ("técnica", "tecnica", "medium", "suporte")):
                        obra.tecnica = limpar(valor)
                    elif any(k in label for k in ("dimen", "medida", "tamanho", "size")):
                        d = padronizar_dimensoes(valor)
                        if d != SEM_INFO:
                            obra.dimensoes = d
                    elif any(k in label for k in ("preço", "valor", "lance", "estimativa")):
                        if obra.preco == SEM_INFO:
                            obra.preco = extrair_preco(valor)
                if obra.ano != SEM_INFO or obra.tecnica != SEM_INFO:
                    break
            except Exception:
                pass

        # ── Descrição ──
        for seletor in [
            ".lot-description",
            ".artwork-description",
            ".obra-descricao",
            ".product-description",
            ".entry-content",
            "[class*='description']",
            "article p",
        ]:
            v = texto_loc(detail_page, seletor)
            if v != SEM_INFO and len(v) > 10:
                obra.descricao = limpar(v[:1000])
                break

        # ── Fallback: varredura geral do texto ──
        if obra.ano == SEM_INFO or obra.tecnica == SEM_INFO or obra.dimensoes == SEM_INFO:
            try:
                texto_geral = detail_page.locator("body").first.inner_text()
                _parse_metadados_texto(texto_geral, obra)
            except Exception:
                pass

    except PlaywrightTimeoutError:
        print(f"  Timeout: {link}", file=sys.stderr)
    except Exception as e:
        print(f"  Erro ao detalhar {link}: {e}", file=sys.stderr)

    return obra


# ─── Descoberta de links ───────────────────────────────────────────────────────

def coletar_links_da_listagem(page: Page) -> list[str]:
    """Extrai links de obras/lotes da página corrente."""
    links: list[str] = []
    seletores = [
        # Padrões específicos de leilão
        ".lot-card a",
        ".lot-item a",
        ".lote-item a",
        ".artwork-card a",
        ".obra-card a",
        # Padrões genéricos
        "[class*='lot'] a[href*='/lote']",
        "[class*='lot'] a[href*='/obra']",
        "[class*='lot'] a[href*='/leilao']",
        "article a[href]",
        ".item a[href]",
    ]
    for seletor in seletores:
        try:
            encontrados = page.locator(seletor).all()
            if encontrados:
                for loc in encontrados:
                    href = loc.get_attribute("href")
                    if href:
                        full = abs_url(href)
                        if full != SEM_INFO and full not in links:
                            links.append(full)
                if links:
                    break
        except Exception:
            pass

    # Fallback: coleta todos os links da página e filtra
    if not links:
        try:
            todos = page.locator("a[href]").all()
            for loc in todos:
                href = loc.get_attribute("href") or ""
                if not href:
                    continue
                full = abs_url(href)
                if (
                    full != SEM_INFO
                    and BASE_URL in full
                    and any(p in full for p in ("/lote/", "/obra/", "/lot/", "/artwork/", "/produto/"))
                    and full not in links
                ):
                    links.append(full)
        except Exception:
            pass

    return links


def scrape_rubem_grilo(
    url: str = ARTISTA_URL,
    max_obras: int = 0,
    headless: bool = True,
) -> pd.DataFrame:
    rows: list[dict] = []
    links_vistos: set[str] = set()

    with sync_playwright() as pw:
        browser: Browser = pw.chromium.launch(
            headless=headless,
            args=[
                "--disable-blink-features=AutomationControlled",
                "--disable-dev-shm-usage",
                "--no-sandbox",
            ],
        )
        context: BrowserContext = browser.new_context(
            viewport={"width": 1440, "height": 900},
            locale="pt-BR",
            user_agent=(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"
            ),
        )
        page = context.new_page()
        detail_page = context.new_page()

        print(f"Abrindo: {url}")
        try:
            page.goto(url, wait_until="domcontentloaded", timeout=60_000)
            page.wait_for_timeout(2_500)
        except Exception as e:
            print(f"Erro ao abrir listagem inicial: {e}", file=sys.stderr)
            context.close()
            browser.close()
            return pd.DataFrame(columns=EXCEL_COLUMNS)

        pagina_num = 1
        parar = False

        while not parar:
            # Aguarda carregamento de cards dinâmicos
            try:
                page.wait_for_selector(
                    ".lot-card, .lot-item, .obra-card, .artwork-card, article, .item",
                    timeout=10_000,
                )
            except PlaywrightTimeoutError:
                pass

            novos_links = coletar_links_da_listagem(page)
            novos_links = [l for l in novos_links if l not in links_vistos]
            print(f"\nPágina {pagina_num} → {len(novos_links)} links novos encontrados")

            if not novos_links:
                print("Nenhum link novo. Encerrando.")
                break

            for link in novos_links:
                if max_obras > 0 and len(rows) >= max_obras:
                    parar = True
                    break
                links_vistos.add(link)
                print(f"  [{len(rows)+1}] Detalhando: {link}")
                obra = extrair_detalhe(detail_page, link)
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
                print(f"    → {obra.titulo} | {obra.tecnica} | {obra.preco}")

            if parar:
                print(f"\nLimite de {max_obras} obras atingido.")
                break

            # Tenta avançar para próxima página
            avancar = False
            for seletor_prox in [
                "a.next-page",
                "a[aria-label='Next']",
                "a[aria-label='Próxima']",
                "button.next-page",
                "button[aria-label*='próximo' i]",
                "a[rel='next']",
                ".pagination a.next",
                ".paginacao a.proximo",
                "li.next a",
            ]:
                try:
                    btn = page.locator(seletor_prox).first
                    if btn.count() > 0 and btn.is_visible():
                        url_antes = page.url
                        btn.click()
                        try:
                            page.wait_for_url(lambda u: u != url_antes, timeout=10_000)
                        except PlaywrightTimeoutError:
                            pass
                        page.wait_for_timeout(2_000)
                        pagina_num += 1
                        avancar = True
                        break
                except Exception:
                    pass

            # Tenta paginação por número
            if not avancar:
                try:
                    prox_num = page.locator(
                        f".pagination a:has-text('{pagina_num + 1}'), "
                        f".paginacao a:has-text('{pagina_num + 1}')"
                    ).first
                    if prox_num.count() > 0:
                        url_antes = page.url
                        prox_num.click()
                        try:
                            page.wait_for_url(lambda u: u != url_antes, timeout=10_000)
                        except PlaywrightTimeoutError:
                            pass
                        page.wait_for_timeout(2_000)
                        pagina_num += 1
                        avancar = True
                except Exception:
                    pass

            if not avancar:
                # Tenta "carregar mais" (infinite scroll / load more)
                try:
                    btn_mais = page.locator(
                        "button:has-text('Ver mais'), button:has-text('Carregar mais'), "
                        "button:has-text('Load More'), a:has-text('Ver mais lotes')"
                    ).first
                    if btn_mais.count() > 0 and btn_mais.is_visible():
                        total_antes = len(links_vistos)
                        btn_mais.click()
                        page.wait_for_timeout(2_500)
                        avancar = True
                        pagina_num += 1
                    else:
                        print("Fim da paginação.")
                        break
                except Exception:
                    print("Fim da paginação.")
                    break

        context.close()
        browser.close()

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
        description=f"Scraper Arremate Arte — {ARTISTA_NOME}",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    parser.add_argument("--url",       default=os.getenv("ARTISTA_URL", ARTISTA_URL))
    parser.add_argument("--max-obras", type=int, default=int(os.getenv("MAX_OBRAS", "0")))
    parser.add_argument("--max-pages", type=int, default=0)
    parser.add_argument("--output",    default=os.getenv("OUTPUT", "artes_RubemGrilo_Arremate.xlsx"))
    parser.add_argument("--headless",  default="true", choices=["true", "false"])
    args = parser.parse_args()

    iniciado = datetime.now()
    df = scrape_rubem_grilo(
        url=args.url,
        max_obras=args.max_obras,
        headless=args.headless.lower() == "true",
    )
    exportar(df, args.output)
    elapsed = datetime.now() - iniciado
    print(f"\nConcluído: {len(df)} obras exportadas para '{args.output}' em {elapsed}.")


if __name__ == "__main__":
    main()
