
# -*- coding: utf-8 -*-
# 4_IOB_ORIENTACOES_FIRE_FOX.py
# Fluxo: Login -> Área temática -> Tributária Federal -> Página Inicial -> Atualizações (Últimas Notícias -> Legislação)
# -> Abrir detalhe de cada item -> Extrair título + texto completo + data -> Consolidar Excel (I:\)

import os
import time
import logging
import re
from pathlib import Path
from datetime import datetime, timedelta
from typing import Optional, List, Dict, Any, Tuple

import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.options import Options as FirefoxOptions
from selenium.webdriver.support.ui import WebDriverWait as W
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, ElementClickInterceptedException
from selenium.webdriver.common.action_chains import ActionChains

# =============== CONFIG ===============
BASE_URL_HOME = "https://www.iobonline.com.br/"
URL_TRIBUTARIA_FEDERAL = "https://www.iobonline.com.br/area/inicial?pagina=tributaria"
WAIT_SEC = 30
HEADLESS = False

# Saída em Excel (I:\)
OUT_DIR = Path(r"I:\\")
OUT_TEMP = OUT_DIR / "Temp_base_atos_extraidos.xlsx"
OUT_BASE = OUT_DIR / "Base_atos_extraidos.xlsx"
OUT_BACKUP = OUT_DIR / f"BACKUP_Base_atos_extraidos_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

# Metadados
FONTE_FIXA = "IOB - FEDERAL"
ESFERA_FIXA = "FEDERAL"
STATUS_CARGA = "novo"

# Layout final
FINAL_COL_ORDER = [
    "Ato", "Descrição", "Esfera", "UF", "Municipio",
    "Data de extração", "Data de publicação", "Fonte", "StatusCarga"
]
DEDUP_COLS = [
    "Ato", "Descrição", "Esfera", "UF", "Municipio", "Data de publicação", "Fonte"
]

DIAS_LIMITE = 5
ENV_PATH = r"C:\Users\a-81006408\PycharmProjects\IOBP.ENV"

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
logger = logging.getLogger(__name__)

# =============== Utils ===============
def ensure_out_dir() -> None:
    OUT_DIR.mkdir(parents=True, exist_ok=True)

def human_sleep(a: float = 0.6, b: float = 1.2) -> None:
    import random
    time.sleep(random.uniform(a, b))

def data_extracao_like_old() -> str:
    """Retorna a data de extração no formato dd/mm/yyyy (ex.: 14/01/2026)."""
    return datetime.today().strftime('%d/%m/%Y')

def normalize_spaces(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "")).strip()

def extract_date_any(text: str) -> str:
    if not text:
        return ""
    m_all = re.findall(r"\d{2}[./]\d{2}[./]\d{4}", text)
    return m_all[-1] if m_all else ""

def to_date_obj(s: str) -> Optional[datetime]:
    if not s:
        return None
    for fmt in ("%d.%m.%Y", "%d/%m/%Y"):
        try:
            return datetime.strptime(s, fmt)
        except Exception:
            pass
    return None

def js_click(driver, elem) -> None:
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", elem)
    driver.execute_script("arguments[0].click();", elem)

# =============== Driver & Login ===============
def build_driver_with_profile(headless: bool = HEADLESS):
    opts = FirefoxOptions()
    if headless:
        opts.add_argument("-headless")
    driver = webdriver.Firefox(options=opts)
    driver.set_page_load_timeout(60)
    return driver

def accept_cookies_if_present(driver) -> None:
    try:
        btns = driver.find_elements(By.CSS_SELECTOR, "#onetrust-accept-btn-handler")
        for b in btns:
            if b.is_displayed():
                b.click()
                logger.info("Cookies aceitos.")
    except Exception:
        pass

def load_env_if_exists(env_path: str) -> Dict[str, str]:
    env: Dict[str, str] = {}
    p = Path(env_path)
    if not p.exists():
        return env
    try:
        text = p.read_text(encoding="utf-8", errors="ignore")
        for raw_line in text.splitlines():
            line = raw_line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, val = line.split("=", 1)
            env[key.strip()] = val.strip().strip("'").strip('"')
    except Exception as e:
        logger.warning("Falha ao ler ENV: %s", e)
    return env

def handle_session_modal(driver) -> None:
    try:
        btn = W(driver, 8).until(
            EC.element_to_be_clickable((By.XPATH, "//span[normalize-space(.)='Encerrar a sessão e logar']"))
        )
        js_click(driver, btn)
        human_sleep(0.6, 1.2)
        logger.info("✅ 'Encerrar a sessão e logar' clicado.")
    except TimeoutException:
        logger.info("Modal de sessão não apareceu.")

def login_iob_simple(driver, user: str, pwd: str) -> bool:
    driver.get(BASE_URL_HOME)
    accept_cookies_if_present(driver)
    human_sleep()
    try:
        login_btn = W(driver, WAIT_SEC).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "a.button.button-login.font-button.enter"))
        )
        js_click(driver, login_btn)
    except TimeoutException:
        logger.error("Botão 'Login' não encontrado.")
        return False
    try:
        W(driver, WAIT_SEC).until(EC.visibility_of_element_located((By.CSS_SELECTOR, "#txtLogin"))).send_keys(user)
        W(driver, WAIT_SEC).until(EC.visibility_of_element_located((By.CSS_SELECTOR, "#txtPassword"))).send_keys(pwd)
        W(driver, WAIT_SEC).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "button.default-btn.light-green-btn.send-login"))
        ).click()
    except TimeoutException:
        logger.error("Falha ao preencher login.")
        return False

    handle_session_modal(driver)

    try:
        W(driver, WAIT_SEC * 2).until(
            EC.any_of(
                EC.url_contains("/area/"),
                EC.url_contains("/home"),
                EC.presence_of_element_located((By.XPATH, "//*[contains(translate(.,'SAIR','sair'),'sair')]"))
            )
        )
        logger.info("✅ Login realizado com sucesso.")
        return True
    except TimeoutException:
        logger.warning("⚠️ Login não confirmado.")
        return False

# =============== Navegação ===============
def open_menu_area_tematica_and_click_tributaria_federal(driver) -> bool:
    logger.info("Abrindo menu 'Área temática' e navegando para 'Tributária Federal'...")
    try:
        menu_li = W(driver, WAIT_SEC).until(
            EC.presence_of_element_located((By.XPATH, "//li[contains(normalize-space(.),'Área temática')]"))
        )
        ActionChains(driver).move_to_element(menu_li).pause(0.2).perform()
        js_click(driver, menu_li)
        human_sleep(0.5, 1.0)
    except TimeoutException:
        logger.warning("Menu não encontrado; usando URL direta.")
        driver.get(URL_TRIBUTARIA_FEDERAL)
        return True

    try:
        link = W(driver, WAIT_SEC).until(
            EC.element_to_be_clickable((By.XPATH, "//a[contains(.,'Tributária Federal')]"))
        )
        js_click(driver, link)
        human_sleep(0.8, 1.5)
        return True
    except TimeoutException:
        logger.warning("Não consegui clicar 'Tributária Federal'; usando URL direta.")
        driver.get(URL_TRIBUTARIA_FEDERAL)
        return True

def click_pagina_inicial(driver) -> bool:
    logger.info("Clicando no link 'Página Inicial'...")
    try:
        link = W(driver, WAIT_SEC).until(
            EC.element_to_be_clickable((By.XPATH, "//a[normalize-space(.)='Página Inicial' and contains(@href,'pagina=tributaria')]"))
        )
        js_click(driver, link)
        human_sleep(0.6, 1.2)
        logger.info("✅ Link 'Página Inicial' clicado.")
        return True
    except TimeoutException:
        logger.error("Não encontrei o link 'Página Inicial'.")
        return False

def click_tabs_in_atualizacoes(driver) -> bool:
    logger.info("Ativando abas na seção 'Atualizações'...")
    try:
        nav_ul = W(driver, WAIT_SEC).until(
            EC.presence_of_element_located((By.XPATH, "//ul[contains(@class,'atualizacoes__navegacao')]"))
        )
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", nav_ul)
        human_sleep(0.3, 0.6)
    except TimeoutException:
        logger.error("Não encontrei a barra de navegação das abas.")
        return False

    # Últimas Notícias
    try:
        noticias_tab = driver.find_element(By.XPATH, "//ul[contains(@class,'atualizacoes__navegacao')]//li[normalize-space(.)='Últimas Notícias']")
        if "activated" not in noticias_tab.get_attribute("class"):
            js_click(driver, noticias_tab)
            human_sleep(0.6, 1.0)
            logger.info("✅ Aba 'Últimas Notícias' ativada.")
    except Exception:
        logger.warning("Não consegui clicar em 'Últimas Notícias'.")

    # Legislação
    try:
        tab = W(driver, WAIT_SEC).until(
            EC.element_to_be_clickable((By.XPATH, "//ul[contains(@class,'atualizacoes__navegacao')]//li[normalize-space(.)='Legislação']"))
        )
        js_click(driver, tab)
        human_sleep(0.6, 1.2)
        W(driver, WAIT_SEC).until(
            EC.presence_of_element_located((By.XPATH, "//section[contains(@class,'atualizacoes__conteudo')]"))
        )
        logger.info("✅ Aba 'Legislação' ativada.")
        return True
    except TimeoutException:
        logger.error("Não consegui ativar a aba 'Legislação'.")
        return False

# =============== Coleta (lista) ===============
def collect_legislacao_nodes(driver) -> List[Any]:
    try:
        return driver.find_elements(
            By.XPATH,
            "//section[contains(@class,'atualizacoes__conteudo')]//article"
            " | //section[contains(@class,'atualizacoes__conteudo')]//li[.//a or .//h3 or .//p]"
        )
    except Exception:
        return []

def find_detail_link(node) -> Optional[str]:
    """Retorna href para detalhe, se existir."""
    try:
        # Link típico de documento: /documento/doc/Documento... ou /documento/doc/...
        a = node.find_element(By.XPATH, ".//a[contains(@href,'/documento/doc/')]")
        href = a.get_attribute("href") or ""
        return href if href else None
    except Exception:
        return None

# =============== Coleta (detalhe) ===============
def extract_detail_fulltext(driver, url: str) -> Tuple[Optional[str], Optional[str], Optional[str]]:
    """
    Abre detalhe em nova aba e extrai:
    - Título (h1)
    - Texto completo (parágrafos com class 'noticia-texto')
    - Data de publicação ("Publicada em dd.mm.yyyy")
    """
    original = driver.current_window_handle
    driver.execute_script("window.open(arguments[0], '_blank');", url)
    W(driver, WAIT_SEC).until(EC.number_of_windows_to_be(2))
    new_tab = [h for h in driver.window_handles if h != original][0]
    driver.switch_to.window(new_tab)

    try:
        # Espera título ou conteúdo
        W(driver, WAIT_SEC).until(
            EC.any_of(
                EC.presence_of_element_located((By.XPATH, "//h1")),
                EC.presence_of_element_located((By.XPATH, "//p[contains(@class,'noticia-texto')]"))
            )
        )

        # Título
        title = None
        for xp in ["//h1", "//article//h1", "//div[contains(@class,'content')]//h1"]:
            try:
                el = driver.find_element(By.XPATH, xp)
                title = normalize_spaces(el.text)
                if title:
                    break
            except Exception:
                pass

        # Data de publicação: "Publicada em dd.mm.yyyy"
        pub_date = None
        try:
            el_date = driver.find_element(By.XPATH, "//*[contains(normalize-space(.),'Publicada em')]")
            pub_date = extract_date_any(el_date.text)
        except Exception:
            # Fallback: tentar na página inteira
            pub_date = extract_date_any(driver.page_source)

        # Texto completo: todos <p class="noticia-texto">
        paragraphs: List[str] = []
        try:
            ps = driver.find_elements(By.XPATH, "//p[contains(@class,'noticia-texto')]")
            for p in ps:
                txt = normalize_spaces(p.text)
                if txt:
                    paragraphs.append(txt)
        except Exception:
            paragraphs = []

        # Fallback: primeiros parágrafos da área de documento
        if not paragraphs:
            try:
                ps = driver.find_elements(By.XPATH, "//div[@id='js-document']//p | //div[contains(@class,'document')]//p")
                for p in ps[:6]:  # limite defensivo
                    txt = normalize_spaces(p.text)
                    if txt:
                        paragraphs.append(txt)
            except Exception:
                pass

        full_body = "\n\n".join(paragraphs) if paragraphs else None
        return title, full_body, pub_date

    finally:
        try:
            driver.close()
        except Exception:
            pass
        driver.switch_to.window(original)

# =============== Extração principal da aba Legislação ===============
def extract_from_legislacao_tab(driver, dias_limite: int = DIAS_LIMITE) -> List[Dict[str, Any]]:
    items: List[Dict[str, Any]] = []
    limite_data = datetime.today() - timedelta(days=dias_limite)

    for _ in range(10):  # paginação defensiva
        nodes = collect_legislacao_nodes(driver)
        if not nodes:
            logger.info("Nenhum card encontrado na aba 'Legislação'.")
            break

        stop_pagination = False

        for n in nodes:
            # Tenta abrir detalhe
            href = find_detail_link(n)
            if href:
                try:
                    title, full_body, pub_date = extract_detail_fulltext(driver, href)
                except Exception as e:
                    logger.debug("Falha ao abrir detalhe: %s", e)
                    title, full_body, pub_date = None, None, None
            else:
                # Fallback: raspa do card
                title, full_body, pub_date = None, None, None
                # Título
                for xp in [".//h3", ".//h2", ".//a[1]", ".//strong[1]"]:
                    try:
                        el = n.find_element(By.XPATH, xp)
                        title = normalize_spaces(el.text)
                        if title:
                            break
                    except Exception:
                        pass
                # Descrição (primeiro parágrafo)
                for xp in [".//p[1]", ".//div[contains(@class,'resumo') or contains(@class,'snippet')][1]"]:
                    try:
                        el = n.find_element(By.XPATH, xp)
                        full_body = normalize_spaces(el.text)
                        if full_body:
                            break
                    except Exception:
                        pass
                pub_date = extract_date_any(title) if title else None

            # Monta item, se houver título ou descrição
            if not title and not full_body:
                continue

            d_obj = to_date_obj(pub_date) if pub_date else None
            if d_obj and d_obj < limite_data:
                stop_pagination = True

            item = {
                "Ato": title or "",
                "Descrição": full_body or "",
                "Esfera": ESFERA_FIXA,
                "UF": "FEDERAL",
                "Municipio": "",
                "Data de extração": data_extracao_like_old(),  # agora em dd/mm/yyyy
                "Data de publicação": pub_date or "",
                "Fonte": FONTE_FIXA,
                "StatusCarga": STATUS_CARGA,
            }
            items.append(item)

        if stop_pagination:
            logger.info("Parando por limite de datas (%d dias).", dias_limite)
            break

        # Tentar botão 'Próximo' da aba
        try:
            nxt = driver.find_element(By.XPATH, "//a[contains(.,'Próximo') or contains(.,'Proximo')]")
            if nxt.is_displayed() and nxt.is_enabled():
                js_click(driver, nxt)
                human_sleep(0.8, 1.5)
            else:
                break
        except Exception:
            break

    logger.info("Itens coletados (Legislação): %d", len(items))
    return items

# =============== Excel ===============
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

    for col in FINAL_COL_ORDER:
        if col not in df.columns:
            df[col] = ""
    df = df[FINAL_COL_ORDER]

    try:
        df.to_excel(OUT_TEMP, sheet_name="dados", index=False)
        logger.info("Temp salvo: %s", OUT_TEMP)
    except Exception as e:
        logger.error("Falha ao salvar temp: %s", e)

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

        # Dedupe primário preserva registros antigos (base + novos)
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

# =============== MAIN ===============
def main() -> None:
    env = load_env_if_exists(ENV_PATH) or {}
    user = env.get("USER_OR") or os.environ.get("USER_OR")
    pwd = env.get("PWD_OR") or os.environ.get("PWD_OR")

    if not user:
        user = input("USER_OR: ").strip()
    if not pwd:
        pwd = input("PWD_OR: ").strip()

    driver = None
    try:
        driver = build_driver_with_profile(HEADLESS)

        if not login_iob_simple(driver, user, pwd):
            return
        if not open_menu_area_tematica_and_click_tributaria_federal(driver):
            return
        if not click_pagina_inicial(driver):
            return
        if not click_tabs_in_atualizacoes(driver):
            return

        items = extract_from_legislacao_tab(driver, dias_limite=DIAS_LIMITE)
        consolidate_to_excel(items)

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
