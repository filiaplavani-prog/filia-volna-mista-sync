# Filia – volná místa

Automatická veřejná stránka obsazenosti kurzů.

- Zdrojový Excel s osobními údaji se do GitHubu **nikdy neukládá**.
- Workflow stáhne soubor z OneDrivu pouze do dočasného runneru.
- `scripts/build_public.py` čte výhradně list `WEB_DATA` a sloupce A:G.
- Na web se publikuje pouze den, čas a text dostupnosti.
- Synchronizace je naplánována na každých 5 minut.

Veřejná stránka je v adresáři `docs/`.
