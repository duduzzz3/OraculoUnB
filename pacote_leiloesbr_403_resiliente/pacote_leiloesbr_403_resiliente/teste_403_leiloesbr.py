from __future__ import annotations

import sys
import textwrap
from pathlib import Path

import requests

URL = "https://leiloesbr.com.br/busca_andamento.asp?pesquisa=&op=2&v=126&tp=%7C50696E74757261732065204772617675726173%7C&b=0&pag=1"
BASE = "https://leiloesbr.com.br/default.asp"

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8",
    "Accept-Language": "pt-BR,pt;q=0.9,en-US;q=0.8,en;q=0.7",
    "Referer": "https://leiloesbr.com.br/",
    "Upgrade-Insecure-Requests": "1",
}


def teste_requests() -> None:
    print("=== Teste 1: requests com sessão e headers ===")
    s = requests.Session()
    s.headers.update(HEADERS)
    try:
        r0 = s.get(BASE, timeout=20)
        print("Homepage:", r0.status_code, "cookies:", len(s.cookies))
    except Exception as e:
        print("Homepage erro:", repr(e))
    try:
        r = s.get(URL, timeout=30)
        print("Busca:", r.status_code)
        print("Final URL:", r.url)
        print("Tamanho HTML:", len(r.text or ""))
        print("Trecho:", (r.text or "")[:300].replace("\n", " "))
        if r.status_code == 403:
            print("Resultado: 403. O ambiente/IP foi recusado pelo servidor.")
        elif "abre_catalogo" in r.text:
            print("Resultado: OK. A listagem contém links de lote.")
        else:
            print("Resultado: resposta recebida, mas sem links esperados.")
    except Exception as e:
        print("Busca erro:", repr(e))


def teste_playwright() -> None:
    print("\n=== Teste 2: Playwright/browser ===")
    try:
        from playwright.sync_api import sync_playwright
    except Exception:
        print("Playwright não está instalado neste ambiente.")
        return
    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True, args=["--no-sandbox"])
            ctx = browser.new_context(
                locale="pt-BR",
                viewport={"width": 1366, "height": 768},
                user_agent=HEADERS["User-Agent"],
            )
            page = ctx.new_page()
            page.goto(BASE, wait_until="domcontentloaded", timeout=45000)
            page.wait_for_timeout(1000)
            resp = page.goto(URL, wait_until="domcontentloaded", timeout=45000)
            page.wait_for_timeout(1500)
            html = page.content()
            status = resp.status if resp else "sem status"
            print("Status:", status)
            print("Tamanho HTML:", len(html))
            print("Contém abre_catalogo:", "abre_catalogo" in html)
            Path("debug_leiloesbr_playwright.html").write_text(html, encoding="utf-8")
            print("HTML salvo em: debug_leiloesbr_playwright.html")
            ctx.close()
            browser.close()
    except Exception as e:
        print("Playwright erro:", repr(e))


if __name__ == "__main__":
    teste_requests()
    teste_playwright()
    print("\nSe requests e Playwright retornarem 403 no Streamlit, rode o scraper localmente ou use o modo --html-file/--html-dir com HTML salvo pelo navegador.")
