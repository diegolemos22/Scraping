
# -*- coding: utf-8 -*-
"""
Scraper SPED com Firefox, salvando no layout IOB.
"""

import os
import time
import pandas as pd
from datetime import datetime, timedelta
from pathlib import Path
from selenium import webdriver
from selenium.webdriver.firefox.options import Options
from bs4 import BeautifulSoup as bs

# ========================= CONFIGURAÇÕES =========================
HEADLESS = False
FIREFOX_PROFILE_PATH = r"C:\\Users\\a-81006408\\AppData\\Roaming\\Mozilla\\Firefox\\Profiles\\77teivsv.default"

# Diretórios (mesmos do IOB)
OUT_DIR = Path(r"C:\\Users\\a-81006408\\OneDrive - Vale S.A\\Documentos\\Dados Controles\\IOB")
OUT_TEMP = OUT_DIR / "Temp_base_atos_extraidos.xlsx"
OUT_BASE = OUT_DIR / "Base_atos_extraidos.xlsx"
OUT_BACKUP = OUT_DIR / "BACKUP_Base_atos_extraidos.xlsx"

FINAL_COL_ORDER = [
    "Ato", "Descrição", "Esfera", "UF", "Municipio",
    "Data de extração", "Data de publicação", "Fonte", "StatusCarga"
]

STATUS_CARGA_FIXO = "novo"

# URLs SPED
PAGE_DESTAQUES = "http://sped.rfb.gov.br/destaques/show/7"
PAGE_BASE = "http://sped.rfb.gov.br/pagina/show/"

# ========================= FUNÇÕES =========================
def build_firefox(headless=False):
    opts = Options()
    if headless:
        opts.add_argument("-headless")
    driver = webdriver.Firefox(options=opts)
    driver.set_page_load_timeout(60)
    return driver

def ensure_out_dir():
    OUT_DIR.mkdir(parents=True, exist_ok=True)

def dedupe_and_save(df):
    for col in FINAL_COL_ORDER:
        if col not in df.columns:
            df[col] = ""
    df = df[FINAL_COL_ORDER]

    df.to_excel(OUT_TEMP, sheet_name="dados", index=False)

    if OUT_BASE.exists():
        df_base = pd.read_excel(OUT_BASE, engine="openpyxl")
        for col in FINAL_COL_ORDER:
            if col not in df_base.columns:
                df_base[col] = ""
        df_all = pd.concat([df_base, df], ignore_index=True)
    else:
        df_all = df.copy()

    df_all = df_all.drop_duplicates(subset=["Ato", "Descrição", "Fonte"], keep="first")

    df_all.to_excel(OUT_BASE, sheet_name="dados", index=False)
    df_all.to_excel(OUT_BACKUP, sheet_name="dados", index=False)

    print(f"✅ Base consolidada salva em {OUT_BASE} (Total: {len(df_all)} registros)")

# ========================= SCRAPER =========================
def scrape_sped():
    items = []
    data_extracao = datetime.today().strftime("%Y-%m-%d")
    data_corte = (datetime.today() - timedelta(days=90)).date()

    driver = build_firefox(HEADLESS)
    driver.get(PAGE_DESTAQUES)
    time.sleep(3)
    soup = bs(driver.page_source, "html.parser")
    driver.quit()

    artigos = soup.find("h2", {"class": "titulo-destaque-ano"})
    if not artigos:
        print("⚠ Nenhum destaque encontrado.")
        return items

    lista_li = artigos.find_all_next("li")
    for li in lista_li:
        try:
            href = li.find("a").get("href")
            pagina = href.split("/")[-1]
            data_txt = li.text.split("\n")[2].strip()[1:-1]
            data_pub = datetime.strptime(data_txt, "%d/%m/%Y").date()
        except Exception:
            continue

        if data_pub > data_corte:
            driver = build_firefox(HEADLESS)
            driver.get(PAGE_BASE + pagina)
            time.sleep(2)
            page_soup = bs(driver.page_source, "html.parser")
            driver.quit()

            try:
                ato = page_soup.find("h1", {"class": "destacado"}).text.strip()
                descricao = page_soup.find(id="conteudo-pagina").text.strip()
                data_publicacao = page_soup.find("article", {"class": "container-conteudo-item grid-3"}).find_all("p")[0].text[-10:]
            except Exception:
                continue

            item = {
                "Ato": ato,
                "Descrição": descricao,
                "Esfera": "FEDERAL",
                "UF": "FEDERAL",
                "Municipio": "",
                "Data de extração": data_extracao,
                "Data de publicação": data_publicacao,
                "Fonte": "SPED",
                "StatusCarga": STATUS_CARGA_FIXO
            }
            items.append(item)

    return items

# ========================= MAIN =========================
def main():
    ensure_out_dir()
    items = scrape_sped()
    if not items:
        print("⚠ Nenhum item encontrado.")
        return
    df = pd.DataFrame(items)
    dedupe_and_save(df)

if __name__ == "__main__":
    main()
