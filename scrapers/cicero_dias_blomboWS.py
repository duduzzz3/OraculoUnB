"""
Scraper Blombo – Cícero Dias
Coleta obras de https://blombo.com/artistas/cicero-dias

Uso básico:
    python cicero_dias_blomboWS.py
    python cicero_dias_blomboWS.py --max-obras 50 --output saida.xlsx --headless true
"""
from __future__ import annotations

import argparse
import os
import re
import sys
from datetime import datetime
from typing import Any

import pandas as pd
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError

# ─── Constantes ───────────────────────────────────────────────────────────────
BASE_URL    = "https://blombo.com"
ARTISTA_URL = f"{BASE_URL}/artistas/cicero-dias"
ARTISTA_NOME = "Cícero Dias"
SEM_INFO    = "Sem informação"

EXCEL_COLUMNS = [
    "Título", "Autor", "Ano", "Técnica", "Dimensões",
    "Preço", "Descrição", "Link da obra", "Link da imagem da obra",
]


# ─── Utilitários ──────────────────────────────────────────────────────────────

def limpar(valor: Any) -> str:
    """Remove espaços redundantes e retorna SEM_INFO se vazio."""
    if valor is None:
        return SEM_INFO
    t = str(valor).replace("\xa0", " ").replace("\u200b", "").strip()
    t = re.sub(r"\s+", " ", t)
    return t if t else SEM_INFO


def padronizar_dimensoes(texto: str) -> str:
    """Converte dimensões para o padrão 'A x B cm' ou 'A x B x C cm'."""
    if not texto or not str(texto).strip():
        return SEM_INFO
    nums = re.findall(r"\d+(?:[.,]\d+)?", str(texto))
    partes = []
    for n in nums[:3]:
        try:
            v = float(n.replace(",", "."))
            partes.append(str(int(v)) if v.is_integer() else str(v).rstrip("0").rstrip("."))
        except ValueError:
            pass
    if len(partes) >= 2:
        return " x ".join(partes) + " cm"
    return SEM_INFO


def texto_ou_sem_info(locator) -> str:
    """Extrai texto de um locator de forma segura."""
    try:
        if locator.count() > 0:
            t = locator.first.inner_text().strip()
            return t if t else SEM_INFO
    except Exception:
        pass
    return SEM_INFO


def extrair_preco(arte) -> float | str:
    """Tenta extrair preço em BRL do card da obra."""
    seletores = [
        "p.special-price span.price",
        "span.price[id^='product-price-']",
        "span.price:not([id^='old-price-'])",
    ]
    for seletor in seletores:
        try:
            loc = arte.locator(seletor).first
            if loc.count() > 0:
                txt = loc.inner_text().strip()
                if txt:
                    valor = txt.replace("R$", "").strip().replace(".", "").replace(",", ".")
                    return round(float(valor), 2)
        except Exception:
            pass
    return SEM_INFO


def extrair_melhor_imagem(arte) -> str:
    """Tenta extrair a URL da imagem de maior qualidade do card."""
    try:
        img = arte.locator("img").first
        # Prefere srcset (maior resolução)
        srcset = img.get_attribute("srcset") or ""
        if srcset:
            partes = [p.strip().split(" ")[0] for p in srcset.split(",") if p.strip()]
            if partes:
                url = partes[-1]
                return ("https:" + url) if url.startswith("//") else url
        for attr in ("src", "data-src", "data-original-src", "data-image"):
            url = img.get_attribute(attr)
            if url:
                return ("https:" + url) if url.startswith("//") else url
    except Exception:
        pass
    return SEM_INFO


# ─── Scraping ─────────────────────────────────────────────────────────────────

def scrape_cicero_dias(
    url: str = ARTISTA_URL,
    max_obras: int = 0,
    headless: bool = True,
) -> pd.DataFrame:
    """Percorre a página do artista no Blombo e coleta todos os dados."""
    rows: list[dict] = []
    links_vistos: set[str] = set()
    parar = False

    with sync_playwright() as pw:
        browser = pw.chromium.launch(
            headless=headless,
            args=["--disable-blink-features=AutomationControlled", "--no-sandbox"],
        )
        context = browser.new_context(
            viewport={"width": 1440, "height": 900},
            user_agent=(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"
            ),
        )
        page = context.new_page()
        detail_page = context.new_page()

        print(f"Abrindo: {url}")
        page.goto(url, wait_until="domcontentloaded", timeout=60_000)
        page.wait_for_timeout(2_000)

        pagina_num = 1
        while not parar:
            artes = page.locator("div[data-product-id]")
            total = artes.count()
            print(f"\nPágina {pagina_num} → {total} obras encontradas")

            if total == 0:
                # Fallback: tenta seletor mais genérico
                artes = page.locator(".products-list .product, .grid-product, .c-product-card")
                total = artes.count()
                if total == 0:
                    print("Nenhuma obra encontrada. Encerrando.")
                    break

            for i in range(total):
                if max_obras > 0 and len(rows) >= max_obras:
                    parar = True
                    break

                arte = artes.nth(i)

                # Link da obra
                link_arte = SEM_INFO
                for seletor_link in ["a.product-image", "a.product__link", "a[href*='/obras/']"]:
                    try:
                        href = arte.locator(seletor_link).first.get_attribute("href")
                        if href:
                            link_arte = href if href.startswith("http") else BASE_URL + href
                            break
                    except Exception:
                        pass

                if link_arte in links_vistos:
                    continue
                links_vistos.add(link_arte)

                # Dados do card
                titulo   = texto_ou_sem_info(arte.locator(".product__name"))
                tecnica  = texto_ou_sem_info(arte.locator(".product__technique"))
                dimensoes = padronizar_dimensoes(
                    texto_ou_sem_info(arte.locator(".product__dimensions"))
                )

                ano = SEM_INFO
                try:
                    ano_loc = arte.locator("span.product__technique + span")
                    if ano_loc.count():
                        raw = ano_loc.first.inner_text().strip().lstrip(", ").strip()
                        m = re.search(r"\b(19|20)\d{2}\b", raw)
                        ano = m.group(0) if m else (raw if raw else SEM_INFO)
                except Exception:
                    pass

                link_imagem = extrair_melhor_imagem(arte)
                preco       = extrair_preco(arte)
                descricao   = SEM_INFO

                # Abre página interna para descrição
                if link_arte != SEM_INFO:
                    try:
                        detail_page.goto(link_arte, wait_until="domcontentloaded", timeout=30_000)
                        detail_page.wait_for_timeout(800)
                        for seletor_desc in [
                            "p.c-product__artist",
                            ".product-description p",
                            ".c-product__description",
                            "[class*='description'] p",
                        ]:
                            loc = detail_page.locator(seletor_desc).first
                            if loc.count():
                                d = loc.inner_text().strip()
                                if d:
                                    descricao = d
                                    break
                        # Melhorar imagem se necessário
                        if link_imagem == SEM_INFO:
                            for seletor_img in [
                                "meta[property='og:image']",
                                ".c-product__main-image img",
                                ".product-image img",
                            ]:
                                loc_img = detail_page.locator(seletor_img).first
                                if loc_img.count():
                                    src = (
                                        loc_img.get_attribute("content")
                                        or loc_img.get_attribute("src")
                                        or loc_img.get_attribute("data-src")
                                    )
                                    if src:
                                        link_imagem = ("https:" + src) if src.startswith("//") else src
                                        break
                    except PlaywrightTimeoutError:
                        print(f"  Timeout ao abrir detalhes: {link_arte}", file=sys.stderr)
                    except Exception as e:
                        print(f"  Erro ao abrir detalhes {link_arte}: {e}", file=sys.stderr)

                row = {
                    "Título":                 titulo,
                    "Autor":                  ARTISTA_NOME,
                    "Ano":                    ano,
                    "Técnica":                tecnica,
                    "Dimensões":              dimensoes,
                    "Preço":                  preco,
                    "Descrição":              descricao,
                    "Link da obra":           link_arte,
                    "Link da imagem da obra": link_imagem,
                }
                rows.append(row)
                print(f"  [{len(rows):>3}] {titulo} | {preco}")

            if parar:
                print(f"\nLimite de {max_obras} obras atingido.")
                break

            # Tenta avançar para a próxima página
            try:
                prox = page.locator(
                    "div.pages a.next.i-next[title='Próximo'], "
                    "a.next[rel='next'], "
                    ".pagination a[aria-label='Next'], "
                    "a[aria-label='Próxima página']"
                ).first
                if prox.count() == 0:
                    print("Fim da paginação.")
                    break
                url_antes = page.url
                prox.click()
                try:
                    page.wait_for_url(lambda u: u != url_antes, timeout=10_000)
                except PlaywrightTimeoutError:
                    pass
                page.wait_for_timeout(2_000)
                pagina_num += 1
            except Exception as e:
                print(f"Sem próxima página ({e}). Encerrando.")
                break

        context.close()
        browser.close()

    df = pd.DataFrame(rows, columns=EXCEL_COLUMNS)
    if not df.empty:
        df = df.drop_duplicates(subset=["Link da obra"], keep="first")
    return df


# ─── Exportação ───────────────────────────────────────────────────────────────

def exportar(df: pd.DataFrame, output: str) -> None:
    """Salva o DataFrame em Excel com formatação básica."""
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
        description=f"Scraper Blombo — {ARTISTA_NOME}",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    parser.add_argument(
        "--url",
        default=os.getenv("ARTISTA_URL", ARTISTA_URL),
        help="URL da página do artista no Blombo",
    )
    parser.add_argument(
        "--max-obras",
        type=int,
        default=int(os.getenv("MAX_OBRAS", "0")),
        help="Limite de obras. 0 = sem limite",
    )
    parser.add_argument(
        "--max-pages",
        type=int,
        default=0,
        help="(ignorado; mantido para compatibilidade com o runner do app.py)",
    )
    parser.add_argument(
        "--output",
        default=os.getenv("OUTPUT", "artes_CiceroDias_Blombo.xlsx"),
        help="Arquivo Excel de saída",
    )
    parser.add_argument(
        "--headless",
        default="true",
        choices=["true", "false"],
        help="Rodar navegador invisível",
    )
    args = parser.parse_args()

    iniciado = datetime.now()
    df = scrape_cicero_dias(
        url=args.url,
        max_obras=args.max_obras,
        headless=args.headless.lower() == "true",
    )
    exportar(df, args.output)
    elapsed = datetime.now() - iniciado
    print(f"\nConcluído: {len(df)} obras exportadas para '{args.output}' em {elapsed}.")


if __name__ == "__main__":
    main()
