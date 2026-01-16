
# -*- coding: utf-8 -*-
# svrs_bpe_mdfe_refatorado_v3.py
# Ajustes:
# - Fonte: salva apenas o path da URL (ex.: "Bpe/Avisos")
# - Datas: "Data de extração" e "Data de publicação" em dd/mm/yyyy

import re, time
from pathlib import Path
from typing import List, Dict, Any, Tuple
from datetime import datetime
from urllib.parse import urlparse

import pandas as pd
from bs4 import BeautifulSoup as BS
from selenium import webdriver
from selenium.webdriver.firefox.options import Options as FirefoxOptions
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait as W
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, WebDriverException

# ================= CONFIG =================
WAIT_SEC = 20
HEADLESS = False  # True para execução sem UI
MAX_ITEMS_PER_PAGE = 300

OUT_DIR = Path("I:\\")  # mantém os caminhos do BPE_NOVO
OUT_TEMP = OUT_DIR / "Temp_base_atos_extraidos.xlsx"
OUT_BASE = OUT_DIR / "Base_atos_extraidos.xlsx"
OUT_BACKUP= OUT_DIR / "BACKUP_Base_atos_extraidos.xlsx"

ESFERA_FIXA = "FEDERAL"
UF_FIXO = "FEDERAL"
MUNICIPIO_FIXO = "FEDERAL"  # se preferir vazio, mude para "" (string vazia)
STATUS_CARGA_PADRAO = "novo"

FINAL_COL_ORDER = [
    "Ato", "Descrição", "Esfera", "UF", "Municipio",
    "Data de extração", "Data de publicação", "Fonte", "StatusCarga",
]

DEDUP_COLS = [
    "Ato", "Descrição", "Esfera", "UF", "Municipio", "Data de publicação", "Fonte",
]

SITES = [
    "https://dfe-portal.svrs.rs.gov.br/Bpe/Avisos",
    "https://dfe-portal.svrs.rs.gov.br/Bpe/Noticias",
    "https://dfe-portal.svrs.rs.gov.br/Mdfe/Avisos",
    "https://dfe-portal.svrs.rs.gov.br/Mdfe/Noticias",
]
SITES2 = [
    "https://dfe-portal.svrs.rs.gov.br/Bpe/Documentos",
    "https://dfe-portal.svrs.rs.gov.br/Mdfe/Documentos",
    "https://dfe-portal.svrs.rs.gov.br/Mdfe/Legislacao",
    "https://dfe-portal.svrs.rs.gov.br/Bpe/Legislacao",
]

# ================= Utils =================
DATE_RX_ANY = re.compile(r"\b\d{2}[./]\d{2}[./]\d{4}\b")
LEADING_RX = re.compile(r"^\s*(\d{2}[./]\d{2}[./]\d{4})\s*[-–—:]*\s*")  # data no começo (dd/mm/yyyy ou dd.mm.yyyy)
NBSP_RX = re.compile(r"\xa0")
SPACE_RX = re.compile(r"\s+")

def normalize_spaces(s: str) -> str:
    if s is None: return ""
    s = NBSP_RX.sub(" ", s)
    s = SPACE_RX.sub(" ", s).strip()
    return s

def parse_date_str(date_str: str) -> str:
    """
    Aceita 'dd/mm/yyyy' ou 'dd.mm.yyyy' e retorna 'dd/mm/yyyy'.
    """
    date_str = (date_str or "").strip().replace(" ", "")
    for fmt in ("%d/%m/%Y", "%d.%m.%Y"):
        try:
            return datetime.strptime(date_str, fmt).strftime("%d/%m/%Y")
        except Exception:
            pass
    return ""

def today_br() -> str:
    """Retorna a data de hoje em 'dd/mm/yyyy'."""
    return datetime.today().strftime("%d/%m/%Y")

def fonte_from_url(url: str) -> str:
    """
    Converte URL para o texto desejado em 'Fonte', extraindo apenas o path,
    sem a barra inicial, sem query/fragmento.
    Ex.: 'https://dfe-portal.svrs.rs.gov.br/Bpe/Avisos' -> 'Bpe/Avisos'
    """
    try:
        parsed = urlparse(url)
        path = (parsed.path or "").lstrip("/")
        # Normaliza eventuais barras duplicadas e remove trailing slash
        path = re.sub(r"/{2,}", "/", path).rstrip("/")
        return path or parsed.netloc or url
    except Exception:
        return url

# ================= Driver =================
def build_driver(headless: bool = HEADLESS):
    opts = FirefoxOptions()
    if headless: opts.add_argument("-headless")
    try:
        opts.set_preference("dom.webdriver.enabled", False)
        opts.set_preference("useAutomationExtension", False)
        opts.set_preference("privacy.trackingprotection.enabled", False)
        opts.set_preference("general.useragent.override",
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:121.0) Gecko/20100101 Firefox/121.0")
    except Exception:
        pass
    driver = webdriver.Firefox(options=opts)
    driver.set_page_load_timeout(60)
    try:
        driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
    except Exception:
        pass
    return driver

def open_url(driver, url: str) -> bool:
    try:
        driver.get(url)
    except WebDriverException:
        return False
    try:
        W(driver, WAIT_SEC).until(EC.presence_of_all_elements_located((By.XPATH, "//body//*")))
        time.sleep(0.8)
    except TimeoutException:
        pass
    return True

# ================= Lógica LEGADA + correções =================
# No legado:
# - SITES (Avisos/Noticias): data=8, título=11, corpo>=13
# - SITES2 (Documentos/Legislação): data=3, título=6, corpo>=7
def parse_article_html(art_html: str,
                       idx_date: int, idx_title: int, idx_body_start: int) -> Tuple[str, str, str]:
    """
    Extrai Ato, Descrição e Data de publicação de um <article> usando:
    1) Índices do legado sobre o texto.
    2) + Correção: captura Ato via tags (h1/h2/h3/a) e aplica 'leading date fix' na Descrição.
    """
    soup = BS(art_html, "html.parser")
    text = soup.get_text("\n")
    lines = text.split("\n")

    # 1) LÓGICA LEGADA VIA ÍNDICES
    date_txt = lines[idx_date] if len(lines) > idx_date else ""
    title_idx = lines[idx_title] if len(lines) > idx_title else ""
    body_idx = "".join(lines[idx_body_start:]) if len(lines) > idx_body_start else ""

    # 2) CORREÇÃO: ATO via tags (prioriza conteúdo sem data)
    title_tag = ""
    for sel in ["h1", "h2", "h3", "a"]:
        el = soup.find(sel)
        if el and normalize_spaces(el.get_text(" ")):
            title_tag = normalize_spaces(el.get_text(" "))
            break

    # Decide o título final
    title_final = normalize_spaces(title_tag or title_idx)

    # Corpo: use o primeiro <p> ou o body_idx
    body_tag = ""
    p = soup.find("p")
    if p and normalize_spaces(p.get_text(" ")):
        body_tag = normalize_spaces(p.get_text(" "))
    body_final = normalize_spaces(body_tag or body_idx)

    # Leading '>' como no legado
    if ">" in body_final:
        try:
            body_final = body_final.split(">", 1)[1]
        except Exception:
            pass

    # 3) Data: tenta explicitamente tags e regex; fallback = índice do legado
    date_final = ""
    # busca em elementos com classes de data
    for xp in ["time", "span", "div"]:
        for el in soup.find_all(xp):
            txt = normalize_spaces(el.get_text(" "))
            m = DATE_RX_ANY.findall(txt)
            if m:
                date_final = parse_date_str(m[-1])
            if date_final: break
        if date_final: break

    if not date_final:
        # data pelos índices
        date_final = parse_date_str(date_txt)
    if not date_final:
        # última data que aparecer no texto
        m_all = DATE_RX_ANY.findall(text)
        if m_all: date_final = parse_date_str(m_all[-1])

    # 4) LEADING DATE FIX na Descrição:
    # se a Descrição começar com data e Data de publicação estiver vazia,
    # move a data para Data de publicação e limpa a Descrição
    if body_final:
        mlead = LEADING_RX.match(body_final)
        if mlead:
            dt_lead = parse_date_str(mlead.group(1))
            if dt_lead and not date_final:
                date_final = dt_lead
            body_final = LEADING_RX.sub("", body_final).strip()

    return title_final, body_final, (date_final or "")

def scrape_group_articles(driver, url: str,
                          idx_date: int, idx_title: int, idx_body_start: int) -> List[Dict[str, Any]]:
    items: List[Dict[str, Any]] = []
    if not open_url(driver, url): return items

    html = driver.page_source
    soup = BS(html, "html.parser")
    arts = soup.find_all("article", {"class": "conteudo-lista__item clearfix"}) or soup.find_all("article")

    for art in arts[:MAX_ITEMS_PER_PAGE]:
        ato, desc, dtpub = parse_article_html(str(art), idx_date, idx_title, idx_body_start)
        if not ato and not desc:
            continue

        item = {
            "Ato": ato,
            "Descrição": desc,
            "Esfera": ESFERA_FIXA,
            "UF": UF_FIXO,
            "Municipio": MUNICIPIO_FIXO,  # mantém compatível com seu legado
            "Data de extração": today_br(),
            "Data de publicação": dtpub,  # já vem em dd/mm/yyyy
            "Fonte": fonte_from_url(url),  # apenas "Bpe/Avisos", etc.
            "StatusCarga": STATUS_CARGA_PADRAO,
        }
        items.append(item)

    return items

# ================= Excel =================
def consolidate_to_excel(items: List[Dict[str, Any]]) -> None:
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    if not items:
        df_empty = pd.DataFrame(columns=FINAL_COL_ORDER)
        df_empty.to_excel(OUT_TEMP, sheet_name="dados", index=False)
        df_empty.to_excel(OUT_BASE, sheet_name="dados", index=False)
        try: df_empty.to_excel(OUT_BACKUP, sheet_name="dados", index=False)
        except Exception: pass
        return

    df = pd.DataFrame(items)
    for col in FINAL_COL_ORDER:
        if col not in df.columns: df[col] = ""
    df = df[FINAL_COL_ORDER]

    # 1) Temp
    df.to_excel(OUT_TEMP, sheet_name="dados", index=False)

    # 2) Base + dedupe
    if OUT_BASE.exists():
        df_base = pd.read_excel(OUT_BASE, engine="openpyxl")
        for col in FINAL_COL_ORDER:
            if col not in df_base.columns: df_base[col] = ""
        df_base = df_base[FINAL_COL_ORDER]
        df_all = pd.concat([df_base, df], ignore_index=True)
    else:
        df_all = df.copy()

    # Dedupe primário (como no legado): Ato + Descrição + Fonte
    df_all = df_all.drop_duplicates(subset=["Ato", "Descrição", "Fonte"], keep="first")

    # Dedupe final (limpeza de texto)
    def _clean(s: Any) -> str:
        if pd.isna(s): return ""
        return normalize_spaces(str(s))

    df_cmp = df_all.copy()
    for col in DEDUP_COLS:
        df_cmp[col] = df_cmp[col].apply(_clean)
    df_out = df_cmp.drop_duplicates(subset=DEDUP_COLS, keep="first")

    df_out.to_excel(OUT_BASE, sheet_name="dados", index=False)
    try: df_out.to_excel(OUT_BACKUP, sheet_name="dados", index=False)
    except Exception: pass

# ================= MAIN =================
def main() -> None:
    try:
        driver = build_driver(HEADLESS)
    except WebDriverException as e:
        print(f"Falha ao iniciar Firefox (geckodriver/instalação): {e}")
        return

    all_items: List[Dict[str, Any]] = []
    try:
        # Grupo 1 (Avisos/Noticias) — índices do legado
        for url in SITES:
            all_items.extend(scrape_group_articles(driver, url,
                idx_date=8, idx_title=11, idx_body_start=13))

        # Grupo 2 (Documentos/Legislação) — índices do legado
        for url in SITES2:
            all_items.extend(scrape_group_articles(driver, url,
                idx_date=3, idx_title=6, idx_body_start=7))

        consolidate_to_excel(all_items)

    finally:
        try: driver.quit()
        except Exception: pass

if __name__ == "__main__":
    main()
