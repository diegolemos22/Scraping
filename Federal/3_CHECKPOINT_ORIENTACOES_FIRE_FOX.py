
import os
import re
import time
import logging
import argparse
from pathlib import Path
from datetime import datetime, timedelta
from typing import Optional, List, Dict, Any, Tuple

import pandas as pd
from bs4 import BeautifulSoup as BS
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.options import Options as FirefoxOptions
from selenium.webdriver.support.ui import WebDriverWait as W, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.action_chains import ActionChains

# ========================= CONFIG =========================
BASE_URL_HOME = "https://www.iobonline.com.br/"
URL_TRIBUTARIA_FEDERAL = "https://www.iobonline.com.br/area/inicial?pagina=tributaria"  # fallback
WAIT_SEC = 30
HEADLESS = False

# Saída em Excel (I:\)
OUT_DIR = Path(r"I:\\")
OUT_TEMP = OUT_DIR / "Temp_base_atos_extraidos.xlsx"
OUT_BASE = OUT_DIR / "Base_atos_extraidos.xlsx"
OUT_BACKUP = OUT_DIR / f"BACKUP_Base_atos_extraidos_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

# Checkpoint opcional (CSV)
CHECKPOINT_PATH = OUT_DIR / "checkpoint_estadual.csv"
CHECKPOINT_EVERY = 3  # salva a cada 3 UFs; use 0 para desativar

# Metadados
FONTE_FIXA = "IOB - ESTADUAL"
ESFERA_FIXA = "ESTADUAL"
STATUS_CARGA = "novo"

# Layout final
FINAL_COL_ORDER = [
    "Ato", "Descrição", "Esfera", "UF", "Municipio",
    "Data de extração", "Data de publicação", "Fonte", "StatusCarga"
]

DEDUP_COLS = [
    "Ato", "Descrição", "Esfera", "UF", "Municipio", "Data de publicação", "Fonte"
]

# Janela de coleta (default; pode ser sobrescrito via --dias)
DIAS_LIMITE = 30

# Paginação máxima por UF (default; pode ser sobrescrito via --max-pages)
MAX_PAGES = 10

# Credenciais (ENV)
ENV_PATH = r"C:\Users\a-81006408\PycharmProjects\IOBP.ENV"

# ---------------------- UFs DESEJADAS ----------------------
UFS_DESEJADAS = {
    'Bahia': 'BA', 'Espírito Santo': 'ES', 'Maranhão': 'MA',
    'Mato Grosso do Sul': 'MS', 'Minas Gerais': 'MG', 'Pará': 'PA', 'Pernambuco': 'PE',
    'Mato Grosso': 'MT', 'Rio de Janeiro': 'RJ',
    'Santa Catarina': 'SC', 'São Paulo': 'SP'
}

UF_CODES_SELECT = [sig.lower() for sig in UFS_DESEJADAS.values()]
UF_CODES_OUTPUT = {sig.lower(): sig for sig in UFS_DESEJADAS.values()}

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
logger = logging.getLogger(__name__)

# ========================= Utils =========================
def ensure_out_dir() -> None:
    OUT_DIR.mkdir(parents=True, exist_ok=True)

def human_sleep(a: float = 0.6, b: float = 1.2) -> None:
    import random
    time.sleep(random.uniform(a, b))

def data_extracao_br() -> str:
    """Retorna a data de extração no formato dd/mm/yyyy."""
    return datetime.today().strftime('%d/%m/%Y')

def normalize_spaces(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "")).strip()

DATE_RX_ANY = re.compile(r"\b\d{2}[./]\d{2}[./]\d{4}\b")

def extract_date_any(text: str) -> str:
    """Extrai a última data (dd/mm/yyyy ou dd.mm.yyyy) de um texto."""
    if not text:
        return ""
    m_all = DATE_RX_ANY.findall(text)
    return m_all[-1] if m_all else ""

def normalize_date_br(d: str) -> str:
    d = (d or "").strip().replace(".", "/")
    try:
        return datetime.strptime(d, "%d/%m/%Y").strftime("%d/%m/%Y")
    except Exception:
        return d

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

# Validação de ato
TIPOS_ATO = [
    "Portaria", "Decreto", "Resolução", "Instrução Normativa",
    "Lei", "Lei Complementar", "Comunicado", "Ato", "Convênio",
    "Ajuste", "Protocolo", "SAT", "SEI", "DICAR", "SUFIS"
]
TIPOS_ATO_RX = re.compile(r"\b(" + "|".join(re.escape(t) for t in TIPOS_ATO) + r")\b", re.IGNORECASE)

STOPWORDS_ATO = {"documento", "selecione um estado", "selecione um estado:", "selecione o estado"}

def is_valid_ato(text: str) -> bool:
    """Ato válido: não stopword, contém tipo de ato + (data OU DOE/DOM), tamanho razoável."""
    if not text:
        return False
    t = normalize_spaces(text)
    if t.lower() in STOPWORDS_ATO:
        return False
    if not TIPOS_ATO_RX.search(t):
        return False
    has_date = bool(DATE_RX_ANY.search(t))
    has_doe = ("DOE" in t) or ("DOM" in t)
    if not (has_date or has_doe):
        return False
    if len(t) > 300:  # evita parágrafos gigantes confundidos com título
        return False
    return True

def clean_description_lines(lines: List[str], ato: Optional[str] = None) -> List[str]:
    """Remove linhas de ruído e o próprio ato da descrição."""
    cleaned: List[str] = []
    NOISE_RX = re.compile(
        r"(Favoritar|Imprimir|Compartilhar|Visualizar|Índice|Acesso Rápido|Matérias Federais|"
        r"Selecione um estado|selecione um estado|Fonte:\s*Editorial IOB)",
        re.IGNORECASE
    )
    for ln in lines:
        ln = normalize_spaces(ln)
        if not ln:
            continue
        if NOISE_RX.search(ln):
            continue
        if ato and ln == ato:
            continue
        if ato and ato in ln:
            ln = normalize_spaces(ln.replace(ato, ""))
        if not ln:
            continue
        cleaned.append(ln)
    return cleaned

# ========================= Driver & Login =========================
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

# ========================= Navegação Estadual =========================
def open_area_tematica_and_click_tributaria_estadual(driver) -> bool:
    logger.info("Abrindo menu 'Área temática' -> 'Tributária Estadual'...")
    try:
        menu_li = W(driver, WAIT_SEC).until(
            EC.presence_of_element_located((By.XPATH, "//li[contains(normalize-space(.),'Área temática')]"))
        )
        ActionChains(driver).move_to_element(menu_li).pause(0.2).perform()
        js_click(driver, menu_li)
        human_sleep(0.5, 1.0)
    except TimeoutException:
        logger.warning("Menu 'Área temática' não encontrado; usando URL de fallback.")
        driver.get(URL_TRIBUTARIA_FEDERAL)

    try:
        link = W(driver, WAIT_SEC).until(
            EC.element_to_be_clickable((By.XPATH, "//a[contains(normalize-space(.),'Tributária Estadual')]"))
        )
        js_click(driver, link)
        human_sleep(0.8, 1.5)
        logger.info("✅ 'Tributária Estadual' clicado.")
        return True
    except TimeoutException:
        logger.error("Não consegui clicar em 'Tributária Estadual'.")
        return False

def select_uf_and_open_pagina_inicial(driver, uf_code_lower: str) -> bool:
    code = (uf_code_lower or "").lower().strip()
    logger.info("Selecionando UF no mapa: %s", code.upper())
    try:
        sel = W(driver, WAIT_SEC).until(EC.presence_of_element_located((By.ID, "mapa__estados")))
        Select(sel).select_by_value(code)
        human_sleep(0.3, 0.8)
    except TimeoutException:
        logger.error("Select de UFs (mapa__estados) não encontrado.")
        return False
    except Exception as e:
        logger.error("Falha ao selecionar UF '%s': %s", code, e)
        return False

    try:
        link = W(driver, WAIT_SEC).until(
            EC.element_to_be_clickable((
                By.XPATH,
                f"//a[contains(@href,'/area/inicial') and contains(@href,'esfera=estadual') and contains(@href,'estado={code}')]"
            ))
        )
        js_click(driver, link)
        human_sleep(0.8, 1.5)
        W(driver, WAIT_SEC).until(EC.url_contains("esfera=estadual"))
        W(driver, WAIT_SEC).until(EC.url_contains(f"estado={code}"))
        W(driver, WAIT_SEC).until(
            EC.presence_of_element_located((By.XPATH, "//section[contains(@class,'atualizacoes__conteudo')]"))
        )
        logger.info("✅ 'Página Inicial' da UF %s aberta.", code.upper())
        return True
    except TimeoutException:
        logger.error("Link/URL/Conteúdo da 'Página Inicial' da UF %s não localizado.", code.upper())
        return False

# ========================= Atualizações -> Legislação =========================
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

    for attempt in range(2):
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
            logger.warning("Tentativa %d: não consegui ativar 'Legislação'.", attempt + 1)
            human_sleep(0.5, 1.0)

    logger.error("Falha ao ativar 'Legislação' após retentativas.")
    return False

def collect_legislacao_links(driver) -> List[Tuple[str, str]]:
    """Anchors de documento (evita cabeçalhos/menus)."""
    anchors = driver.find_elements(
        By.XPATH,
        "//section[contains(@class,'atualizacoes__conteudo')]//a[contains(@href,'/documento/doc/')]"
    )
    links: List[Tuple[str, str]] = []
    seen = set()
    for a in anchors:
        try:
            href = a.get_attribute("href") or ""
            if not href or href in seen:
                continue
            text = normalize_spaces(a.text)
            seen.add(href)
            links.append((href, text))
        except Exception:
            continue
    logger.info("Links de documento encontrados: %d", len(links))
    return links

def fallback_card_description_for_link(driver, href: str, ato: Optional[str]) -> str:
    """Fallback de descrição: usa <p> vizinho ao link, limpando ruído e o próprio Ato."""
    try:
        el = driver.find_element(By.XPATH, f"//a[@href='{href}']")
        for xp in ["following::p[1]", "../../p[1]", "../p[1]"]:
            try:
                p = el.find_element(By.XPATH, xp)
                raw = p.text
                lines = clean_description_lines(raw.split("\n"), ato=ato)
                if lines:
                    return " ".join(lines)
            except Exception:
                pass
    except Exception:
        pass
    return ""

def find_ato_strong(driver, page_html: Optional[str] = None) -> Optional[str]:
    """Procura <strong> em containers de documento; escolhe melhor por heurística; valida."""
    candidates: List[str] = []
    xpath_targets = [
        "//div[@id='js-document']//strong",
        "//div[contains(@class,'document')]//strong",
        "//article//strong",
        "//main//strong",
        "//p/strong"
    ]
    for xp in xpath_targets:
        try:
            els = driver.find_elements(By.XPATH, xp)
            for el in els[:20]:
                txt = normalize_spaces(el.text)
                if txt:
                    candidates.append(txt)
        except Exception:
            pass

    if not candidates and page_html:
        try:
            soup = BS(page_html, "html.parser")
            for st in soup.find_all("strong"):
                t = normalize_spaces(st.get_text(" "))
                if t:
                    candidates.append(t)
        except Exception:
            pass

    # filtra e pontua
    valid = [(score_ato_candidate(t), t) for t in candidates if is_valid_ato(t)]
    if not valid:
        return None
    valid.sort(reverse=True)
    return valid[0][1]

def score_ato_candidate(text: str) -> int:
    """Score simples para escolher o melhor <strong> como Ato."""
    if not text:
        return -1
    t = text.strip()
    score = 0
    if TIPOS_ATO_RX.search(t):
        score += 4
    if "DOE" in t or "DOM" in t:
        score += 2
    dates = DATE_RX_ANY.findall(t)
    if len(dates) >= 2:
        score += 3
    elif len(dates) == 1:
        score += 1
    if len(t) <= 220:
        score += 1
    return score

def extract_detail_fulltext(driver, url: str, href_for_fallback: Optional[str] = None) -> Tuple[Optional[str], Optional[str], Optional[str]]:
    """Extrai Ato (<strong> válido), Data e Descrição do detalhe do documento."""
    original = driver.current_window_handle
    driver.execute_script("window.open(arguments[0], '_blank');", url)
    W(driver, WAIT_SEC).until(EC.number_of_windows_to_be(2))
    new_tab = [h for h in driver.window_handles if h != original][0]
    driver.switch_to.window(new_tab)
    try:
        W(driver, WAIT_SEC).until(
            EC.any_of(
                EC.presence_of_element_located((By.XPATH, "//strong")),
                EC.presence_of_element_located((By.XPATH, "//h1")),
                EC.presence_of_element_located((By.XPATH, "//p"))
            )
        )
        page_html = driver.page_source

        # ---- ATO (via <strong> válido) ----
        ato = find_ato_strong(driver, page_html=page_html)

        # ---- DATA ----
        pub_date = None
        for xp in [
            "//*[contains(translate(.,'PUBLICADA','publicada'),'publicada em')]",
            "//*[contains(translate(.,'PUBLICADO','publicado'),'publicado em')]",
            "//*[contains(translate(.,'ATUALIZADA','atualizada'),'atualizada em')]",
            "//*[contains(translate(.,'ATUALIZADO','atualizado'),'atualizado em')]",
        ]:
            try:
                el = driver.find_element(By.XPATH, xp)
                d = extract_date_any(el.text)
                if d:
                    pub_date = d
                    break
            except Exception:
                pass

        if not pub_date:
            try:
                body = driver.find_element(By.XPATH, "//main|//article|//body")
                txt = normalize_spaces(body.text)[:2000]
                pub_date = extract_date_any(txt)
            except Exception:
                pub_date = extract_date_any(page_html)

        # ---- DESCRIÇÃO ----
        body_paras: List[str] = []
        for xp in [
            "//div[@id='js-document']//p",
            "//div[contains(@class,'document')]//p",
            "//article//p"
        ]:
            try:
                ps = driver.find_elements(By.XPATH, xp)
                for p in ps[:12]:
                    txt = normalize_spaces(p.text)
                    if txt:
                        body_paras.append(txt)
                if body_paras:
                    break
            except Exception:
                pass

        # limpeza e remoção de ruído/ato
        lines = clean_description_lines(body_paras, ato=ato)
        full_body = " ".join(lines) if lines else None

        return (ato or None), (full_body or None), (pub_date or None)
    finally:
        try:
            driver.close()
        except Exception:
            pass
        driver.switch_to.window(original)

def extract_from_legislacao_tab(driver, dias_limite: int, uf_sigla: str) -> List[Dict[str, Any]]:
    """Percorre links de documento, extrai detalhe; valida Ato e Descrição; pula itens ruins/fora da janela."""
    items: List[Dict[str, Any]] = []
    limite_data = datetime.today() - timedelta(days=dias_limite)

    for page in range(MAX_PAGES):  # <<<< USANDO MAX_PAGES CONFIGURÁVEL
        links = collect_legislacao_links(driver)
        if not links:
            if page == 0:
                logger.info("Nenhum link de documento encontrado na aba 'Legislação'.")
            break

        for href, link_text in links:
            ato, full_body, pub_date = None, None, None
            try:
                ato, full_body, pub_date = extract_detail_fulltext(driver, href, href_for_fallback=href)
            except Exception as e:
                logger.debug("Falha ao abrir detalhe (%s): %s", href, e)

            # Valida Ato — se não há <strong> válido, tenta o texto do link; se ainda inválido, pula
            if not ato or not is_valid_ato(ato):
                link_clean = normalize_spaces(link_text)
                if is_valid_ato(link_clean):
                    ato = link_clean
                else:
                    logger.debug("Ato inválido/ruído; item descartado. (Ato=%r, link=%r)", ato, link_clean)
                    continue

            pub_date = normalize_date_br(pub_date)

            # Descrição mínima
            if not full_body or len(full_body) < 10:
                fb = fallback_card_description_for_link(driver, href, ato=ato)
                full_body = fb if fb else ""

            # Corte por data — APENAS pula o item
            d_obj = to_date_obj(pub_date) if pub_date else None
            if d_obj and d_obj < limite_data:
                logger.debug(
                    "Item ignorado por estar fora da janela (%s < %s).",
                    pub_date, (datetime.today() - timedelta(days=dias_limite)).strftime('%d/%m/%Y')
                )
                continue

            item = {
                "Ato": ato,
                "Descrição": full_body,
                "Esfera": ESFERA_FIXA,
                "UF": uf_sigla,
                "Municipio": "",
                "Data de extração": data_extracao_br(),
                "Data de publicação": pub_date or "",
                "Fonte": FONTE_FIXA,
                "StatusCarga": STATUS_CARGA,
            }
            items.append(item)

        # Paginação
        try:
            nxt = driver.find_element(By.XPATH, "//a[contains(.,'Próximo') or contains(.,'Proximo')]")
            if nxt.is_displayed() and nxt.is_enabled():
                js_click(driver, nxt)
                human_sleep(0.8, 1.5)
            else:
                break
        except Exception:
            break

    vazias = sum(1 for it in items if not it.get("Descrição"))
    if vazias:
        logger.warning("UF %s: %d itens sem descrição após extração.", uf_sigla, vazias)
    logger.info("Itens coletados (UF=%s): %d", uf_sigla, len(items))
    return items

# ========================= Excel =========================
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

    # Temp
    try:
        df.to_excel(OUT_TEMP, sheet_name="dados", index=False)
        logger.info("Temp salvo: %s (rows=%d)", OUT_TEMP, len(df))
    except Exception as e:
        logger.error("Falha ao salvar temp: %s", e)

    # Base + dedupe
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

        logger.info("Consolidação: antes da dedupe primária (rows=%d)", len(df_all))

        # Dedupe primária AJUSTADA
        df_all = df_all.drop_duplicates(
            subset=["UF", "Ato", "Descrição", "Fonte", "Data de publicação"],
            keep="first"
        )
        logger.info("Após dedupe primária (rows=%d)", len(df_all))

        # Dedupe final (limpeza + normalização de data)
        df_cmp = df_all.copy()
        for col in DEDUP_COLS:
            df_cmp[col] = df_cmp[col].apply(dedupe_text)
        df_cmp["Data de publicação"] = df_cmp["Data de publicação"].apply(normalize_date_br)

        before_final = len(df_cmp)
        df_out = df_cmp.drop_duplicates(subset=DEDUP_COLS, keep="first")
        logger.info("Após dedupe final (rows=%d, removidos=%d)", len(df_out), before_final - len(df_out))

        df_out.to_excel(OUT_BASE, sheet_name="dados", index=False)
        logger.info("Base consolidada: %s (rows=%d)", OUT_BASE, len(df_out))
        try:
            df_out.to_excel(OUT_BACKUP, sheet_name="dados", index=False)
            logger.info("Backup salvo: %s", OUT_BACKUP)
        except Exception as e:
            logger.warning("Falha ao salvar backup: %s", e)
    except Exception as e:
        logger.error("Falha na consolidação da base: %s", e)

# ========================= MAIN =========================
def main() -> None:
    # --- declare 'global' antes de qualquer uso ---
    global DIAS_LIMITE, MAX_PAGES, HEADLESS

    # --- CLI (novos parâmetros) ---
    ap = argparse.ArgumentParser(description="IOB - Tributária Estadual por UF (coleta de Ato/Descrição)")
    ap.add_argument("--dias", type=int, default=DIAS_LIMITE, help="Janela de dias para trás (ex.: 30)")
    ap.add_argument("--max-pages", type=int, default=MAX_PAGES, help="Máximo de páginas por UF (ex.: 50)")
    ap.add_argument("--headless", action="store_true", help="Executa o Firefox em modo headless")
    args = ap.parse_args()

    # aplica CLI sobre as configs globais
    DIAS_LIMITE = int(args.dias)
    MAX_PAGES = int(args.max_pages)
    HEADLESS = bool(args.headless)

    env = load_env_if_exists(ENV_PATH) or {}
    user = env.get("USER_OR") or os.environ.get("USER_OR")
    pwd = env.get("PWD_OR") or os.environ.get("PWD_OR")

    if not user:
        user = input("USER_OR: ").strip()
    if not pwd:
        pwd = input("PWD_OR: ").strip()

    driver = None
    all_items: List[Dict[str, Any]] = []

    try:
        driver = build_driver_with_profile(HEADLESS)

        if not login_iob_simple(driver, user, pwd):
            logger.error("Login falhou. Encerrando.")
            return

        if not open_area_tematica_and_click_tributaria_estadual(driver):
            logger.error("Não foi possível abrir 'Tributária Estadual'. Encerrando.")
            return

        for idx, uf_code_lower in enumerate(UF_CODES_SELECT, start=1):
            uf_sigla = UF_CODES_OUTPUT.get(uf_code_lower, uf_code_lower.upper())
            logger.info("=== Iniciando UF %s (%d/%d) ===", uf_sigla, idx, len(UF_CODES_SELECT))

            if not select_uf_and_open_pagina_inicial(driver, uf_code_lower):
                logger.warning("Pulando UF %s por falha de navegação.", uf_sigla)
                if not open_area_tematica_and_click_tributaria_estadual(driver):
                    logger.error("Falha ao retornar ao seletor estadual. Interrompendo.")
                    break
                continue

            if click_tabs_in_atualizacoes(driver):
                items_uf = extract_from_legislacao_tab(driver, dias_limite=DIAS_LIMITE, uf_sigla=uf_sigla)
                logger.info("UF %s: %d itens extraídos.", uf_sigla, len(items_uf))
                all_items.extend(items_uf)
            else:
                logger.warning("Não consegui ativar abas 'Atualizações' para UF %s.", uf_sigla)

            if CHECKPOINT_EVERY and (idx % CHECKPOINT_EVERY == 0) and all_items:
                try:
                    pd.DataFrame(all_items)[FINAL_COL_ORDER].to_csv(CHECKPOINT_PATH, index=False, encoding="utf-8")
                    logger.info("Checkpoint salvo (%d itens) em: %s", len(all_items), CHECKPOINT_PATH)
                except Exception as e:
                    logger.warning("Falha ao salvar checkpoint: %s", e)

            if not open_area_tematica_and_click_tributaria_estadual(driver):
                logger.error("Falha ao retornar ao seletor estadual após UF %s. Interrompendo.", uf_sigla)
                break

            logger.info("=== Finalizado UF %s ===", uf_sigla)
            human_sleep(0.8, 1.8)

        if all_items:
            logger.info("Execução completa. Consolidando resumo final (%d itens) ...", len(all_items))
            consolidate_to_excel(all_items)
        else:
            logger.warning("Nenhum item foi extraído em toda a execução.")

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
