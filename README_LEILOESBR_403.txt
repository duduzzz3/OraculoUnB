PACOTE LEILÕES BR — VERSÃO RESILIENTE A 403

O que mudou:
1. O scraper continua compatível com o Oráculo Cultural:
   python leiloesbrWS_sem_miniaturas.py --max-obras 100 --output artes_LeiloesBR.xlsx --headless true

2. Ele tenta requests + fallback Playwright, mas agora também tem um modo offline:
   python leiloesbrWS_sem_miniaturas.py --html-file pagina_salva.html --output artes_LeiloesBR.xlsx
   python leiloesbrWS_sem_miniaturas.py --html-dir htmls_leiloesbr --output artes_LeiloesBR.xlsx

3. O modo HTML é o contorno mais confiável quando o Leilões BR bloqueia o IP do Streamlit Cloud.
   Abra a página no seu navegador, salve como HTML e rode o scraper localmente usando --html-file ou --html-dir.
   Depois importe o Excel gerado no Streamlit pela aba Coleta > Importar planilha.

4. Teste de diagnóstico:
   python teste_403_leiloesbr.py

Se o teste der 403 tanto em requests quanto em Playwright no Streamlit Cloud, o bloqueio é do servidor contra aquele ambiente/IP.
Nesse cenário, não há garantia de correção apenas no código do app hospedado no Streamlit.
Use coleta local ou modo HTML salvo.

Instalação:
- Coloque scrapers/leiloesbrWS_sem_miniaturas.py na pasta scrapers do projeto.
- Substitua o app.py se o seu ainda não tiver a opção Leilões BR.
- Suba requirements.txt e packages.txt se ainda não estiverem no GitHub.
