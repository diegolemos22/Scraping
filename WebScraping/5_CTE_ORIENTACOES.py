
# -*- coding: utf-8 -*-
"""
Coleta conteúdos do Portal do CT-e e consolida no Excel padronizado em I:\Base_atos_extraidos.xlsx

Melhorias aplicadas:
- HEADLESS = False por padrão (abre janela); pode usar --headless na CLI para ocultar
- page_load_strategy = "eager" para reduzir timeouts desnecessários
- page load timeout aumentado para 90s
- Navegação resiliente com safe_get() usando window.stop() em caso de TimeoutException
- Espera por elementos de conteúdo principais (tituloConteudo / indentacao*)
"""
import logging
import re
import time
from pathlib import Path
from datetime import datetime
from typing import List, Dict, Any
import argparse

import pandas as pd
from bs4 import BeautifulSoup as bs

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.options import Options as FirefoxOptions
from selenium.webdriver.firefox.service import Service as FirefoxService
from selenium.webdriver.support.ui import WebDriverWait as W
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException

# ================= CONFIG =================
HEADLESS = False  # padrão: abrir a janela; mude via CLI --headless
WAIT_SEC = 30     # espera por presença de elementos principais
PAGE_LOAD_TIMEOUT_SEC = 90

# Diretórios e arquivos (mesma base dos outros robôs)
OUT_DIR = Path(r"I:\\")
OUT_TEMP = OUT_DIR / "Temp_base_atos_extraidos.xlsx"
OUT_BASE = OUT_DIR / "Base_atos_extraidos.xlsx"
OUT_BACKUP = OUT_DIR / f"BACKUP_Base_atos_extraidos_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

# Metadados
FONTE_FIXA = "CTe"
ESFERA_FIXA = "FEDERAL"
UF_FIXA = "FEDERAL"
STATUS_CARGA = "novo"

# Layout final
FINAL_COL_ORDER = [
    "Ato", "Descrição", "Esfera", "UF", "Municipio",
    "Data de extração", "Data de publicação", "Fonte", "StatusCarga"
]

# Dedup final (ignora "Data de extração")
DEDUP_COLS = [
    "Ato", "Descrição", "Esfera", "UF", "Municipio", "Data de publicação", "Fonte"
]

# URLs do portal (mantidas conforme script legado)
CTE_URLS = {
    "Informes": "http://www.cte.fazenda.gov.br/portal/listaConteudo.aspx?tipoConteudo=kMvhq9Tcqys=",
    "Visualizador DF-e": "http://www.cte.fazenda.gov.br/portal/listaConteudo.aspx?tipoConteudo=F+LZ7kF1I4U=",
    "Manuais": "http://www.cte.fazenda.gov.br/portal/listaConteudo.aspx?tipoConteudo=YIi+H8VETH0=",
    "Esquemas XML": "http://www.cte.fazenda.gov.br/portal/listaConteudo.aspx?tipoConteudo=0xlG1bdBass=",
    "Notas Técnicas": "http://www.cte.fazenda.gov.br/portal/listaConteudo.aspx?tipoConteudo=Y0nErnoZpsg=",
    "Diversos": "http://www.cte.fazenda.gov.br/portal/listaConteudo.aspx?tipoConteudo=IV+EmHlFEfQ=",
    # Atenção: "Ajustes SINEF" abaixo está repetindo o mesmo tipoConteudo de "Informes".
    # Se houver URL específica, substitua aqui:
    "Ajustes SINEF": "http://www.cte.fazenda.gov.br/portal/listaConteudo.aspx?tipoConteudo=kMvhq9Tcqys=",
    "Atos COTEPE": "http://www.cte.fazenda.gov.br/portal/listaConteudo.aspx?tipoConteudo=GBUaktfNJMU=",
    "Convênios": "http://www.cte.fazenda.gov.br/portal/listaConteudo.aspx?tipoConteudo=7zEQFBPObw0=",
}

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
logger = logging.getLogger(__name__)


# ================= Utils =================
def ensure_out_dir() -> None:
    OUT_DIR.mkdir(parents=True, exist_ok=True)


def data_extracao_like_old() -> str:
    """Mantém o formato 'YYYY-DD-MM' para compatibilidade."""
    hoje = datetime.today().strftime('%d.%m.%Y')
    return datetime.strptime(hoje, '%d.%m.%Y').strftime('%Y-%d-%m')


def normalize_spaces(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "")).strip()


def extract_date_any(text: str) -> str:
    """Extrai a última data dd.mm.yyyy ou dd/mm/yyyy do texto."""
    if not text:
        return ""
    m_all = re.findall(r"\b\d{2}[./]\d{2}[./]\d{4}\b", text)
    return m_all[-1] if m_all else ""


# ================= Driver =================
def build_driver(headless: bool = HEADLESS):
    opts = FirefoxOptions()

    # Estratégia de carregamento: não esperar recursos tardios (scripts/imagens ao final)
    opts.page_load_strategy = "eager"

    # Proxy do sistema (útil em ambiente corporativo)
    # 5 = proxy herdado do SO; ajuste conforme necessidade
    opts.set_preference("network.proxy.type", 5)

    if headless:
        opts.add_argument("-headless")

    service = FirefoxService()  # geckodriver do PATH
    driver = webdriver.Firefox(options=opts, service=service)

    # Mais folga para páginas lentas
    driver.set_page_load_timeout(PAGE_LOAD_TIMEOUT_SEC)
    return driver


def safe_get(driver, url: str, attempts: int = 2) -> None:
    """
    Navega para a URL com tentativas e fallback de window.stop() para aproveitar o que já carregou.
    """
    last_exc = None
    for i in range(attempts):
        try:
            driver.get(url)
            return
        except TimeoutException as e:
            last_exc = e
            logger.warning("Timeout ao carregar %s (tentativa %d/%d). Tentando window.stop().",
                           url, i + 1, attempts)
            try:
                # Para o carregamento e tenta usar o que já existe na página
                driver.execute_script("window.stop();")
                return
            except Exception:
                time.sleep(2)  # pequena folga e tenta de novo
    # Se chegou aqui, repropaga o último erro
    raise last_exc


def wait_page_ready(driver) -> None:
    """
    Espera pela presença de elementos típicos do conteúdo do portal sem depender do onload completo.
    """
    try:
        W(driver, WAIT_SEC).until(
            EC.presence_of_any_elements_located(
                (By.CSS_SELECTOR, "span.tituloConteudo, div.indentacaoConteudo, div.indentacaoNormal, p")
            )
        )
    except TimeoutException:
        # segue mesmo assim; muitas vezes o HTML já está presente
        pass


# ================= Parse Helpers =================
def parse_cte_page(html: str) -> List[Dict[str, Any]]:
    """
    Extrai itens de uma página do CT-e procurando:
    - títulos: <span class="tituloConteudo">
    - parágrafo mais próximo após o título: <p>
    Também tenta containers alternativos (indentacaoNormal / indentacaoConteudo).
    """
    soup = bs(html, "html.parser")
    items: List[Dict[str, Any]] = []

    # Preferência: todos os títulos na página
    titles = soup.select("span.tituloConteudo")
    if not titles:
        # fallback: procurar por headings/link com aparência de título
        titles = soup.select("div.indentacaoConteudo span, div.indentacaoNormal span")

    for t in titles:
        ato = normalize_spaces(t.get_text())
        if not ato:
            continue
        # primeiro <p> após o título
        p = t.find_next("p")
        corpo = normalize_spaces(p.get_text()) if p else ""
        # remove duplicação do título no corpo
        if corpo and ato in corpo:
            corpo = normalize_spaces(corpo.replace(ato, ""))
        if not corpo:
            corpo = "Sem informação"

        pub_date = extract_date_any(ato) or extract_date_any(corpo) or ""
        item = {
            "Ato": ato,
            "Descrição": corpo,
            "Esfera": ESFERA_FIXA,
            "UF": UF_FIXA,
            "Municipio": "",
            "Data de extração": data_extracao_like_old(),
            "Data de publicação": pub_date,
            "Fonte": FONTE_FIXA,
            "StatusCarga": STATUS_CARGA,
        }
        items.append(item)

    # Fallback adicional: se nada foi extraído, tenta containers comuns
    if not items:
        try:
            blocks = soup.select("div.indentacaoConteudo, div.indentacaoNormal")
            for b in blocks:
                spans = b.select("span")
                ps = b.select("p")
                for i, sp in enumerate(spans):
                    ato = normalize_spaces(sp.get_text())
                    corpo = normalize_spaces(ps[i].get_text()) if i < len(ps) else ""
                    if corpo and ato in corpo:
                        corpo = normalize_spaces(corpo.replace(ato, ""))
                    if not ato and not corpo:
                        continue
                    pub_date = extract_date_any(ato) or extract_date_any(corpo) or ""
                    items.append({
                        "Ato": ato or "",
                        "Descrição": corpo or "Sem informação",
                        "Esfera": ESFERA_FIXA,
                        "UF": UF_FIXA,
                        "Municipio": "",
                        "Data de extração": data_extracao_like_old(),
                        "Data de publicação": pub_date,
                        "Fonte": FONTE_FIXA,
                        "StatusCarga": STATUS_CARGA,
                    })
        except Exception:
            pass

    return items


# ================= Scrape =================
def scrape_cte_section(driver, url: str) -> List[Dict[str, Any]]:
    """Abre a URL, espera carregar e extrai itens via BeautifulSoup (com fallback)."""
    logger.info("Acessando: %s", url)
    safe_get(driver, url)
    wait_page_ready(driver)

    html = driver.page_source
    items = parse_cte_page(html)

    # Fallback por Selenium direto se nada vier
    if not items:
        try:
            spans = driver.find_elements(By.CSS_SELECTOR, "span.tituloConteudo")
            for sp in spans:
                ato = normalize_spaces(sp.text)
                # primeiro p seguinte ao span
                try:
                    p = sp.find_element(By.XPATH, "following::p[1]")
                    corpo = normalize_spaces(p.text)
                except NoSuchElementException:
                    corpo = ""
                if corpo and ato in corpo:
                    corpo = normalize_spaces(corpo.replace(ato, ""))
                if not corpo:
                    corpo = "Sem informação"
                pub_date = extract_date_any(ato) or extract_date_any(corpo) or ""
                items.append({
                    "Ato": ato or "",
                    "Descrição": corpo or "Sem informação",
                    "Esfera": ESFERA_FIXA,
                    "UF": UF_FIXA,
                    "Municipio": "",
                    "Data de extração": data_extracao_like_old(),
                    "Data de publicação": pub_date,
                    "Fonte": FONTE_FIXA,
                    "StatusCarga": STATUS_CARGA,
                })
        except Exception:
            pass

    logger.info("Itens extraídos nesta página: %d", len(items))
    return items


# ================= Excel =================
def dedupe_text(s: str) -> str:
    if pd.isna(s):
        return ""
    return normalize_spaces(str(s))


def consolidate_to_excel(items: List[Dict[str, Any]]) -> None:
    if not items:
        logger.info("Nenhum item para salvar.")
        return

    ensure_out_dir()
    df = pd.DataFrame(items)

    # Garante presença/ordem
    for col in FINAL_COL_ORDER:
        if col not in df.columns:
            df[col] = ""
    df = df[FINAL_COL_ORDER]

    # 1) Temporário
    try:
        df.to_excel(OUT_TEMP, sheet_name="dados", index=False)
        logger.info("Temp salvo: %s", OUT_TEMP)
    except Exception as e:
        logger.error("Falha ao salvar temp: %s", e)

    # 2) Consolidação (append + dedupe)
    try:
        if OUT_BASE.exists():
            df_base = pd.read_excel(OUT_BASE, engine="openpyxl")
            for col in FINAL_COL_ORDER:
                if col not in df_base.columns:
                    df_base[col] = ""
            df_base = df_base[FINAL_COL_ORDER]
            df_all = pd.concat([df_base, df], ignore_index=True)
        else:
            df_all = df.copy()

        # Dedupe primário: mantém o primeiro (da base)
        df_all = df_all.drop_duplicates(subset=["Ato", "Descrição", "Fonte"], keep="first")

        # Dedupe final ignorando "Data de extração"
        df_cmp = df_all.copy()
        for col in DEDUP_COLS:
            df_cmp[col] = df_cmp[col].apply(dedupe_text)
        df_out = df_cmp.drop_duplicates(subset=DEDUP_COLS, keep="first")

        df_out.to_excel(OUT_BASE, sheet_name="dados", index=False)
        logger.info("Base consolidada: %s", OUT_BASE)

        # Backup com timestamp
        try:
            df_out.to_excel(OUT_BACKUP, sheet_name="dados", index=False)
            logger.info("Backup salvo: %s", OUT_BACKUP)
        except Exception as e:
            logger.warning("Falha ao salvar backup: %s", e)

    except Exception as e:
        logger.error("Falha na consolidação da base: %s", e)


# ================= MAIN =================
def main():
    parser = argparse.ArgumentParser(description="Scraper do Portal CT-e (com navegação resiliente).")
    parser.add_argument("--headless", action="store_true", help="Executa sem interface gráfica.")
    args = parser.parse_args()

    driver = None
    try:
        driver = build_driver(headless=args.headless or HEADLESS)
        all_items: List[Dict[str, Any]] = []

        for name, url in CTE_URLS.items():
            try:
                logger.info("Coletando seção: %s", name)
                items = scrape_cte_section(driver, url)
                all_items.extend(items)
                # pequena pausa para não sobrecarregar
                time.sleep(0.6)
            except Exception as e:
                logger.exception("Falha na seção '%s': %s", name, e)

        # Consolida tudo
        consolidate_to_excel(all_items)
    except Exception as e:
        logger.exception("Erro geral: %s", e)
    finally:
        if driver:
            try:
                driver.quit()
            except Exception:
                pass
        logger.info("Driver finalizado.")


if __name__ == "__main__":
    main()
