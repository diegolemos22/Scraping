# scrap_iob_alertas_extracao.py
# Fluxo: Login -> Core Home -> Meu Espaço (menu) -> Meus Alertas
# -> Clicar sino por NOME do alerta -> Ver detalhes (data alvo)
# -> Extrair itens (blocos municipais) -> Consolidar no Excel (novo layout)

import os
import time
import logging
from pathlib import Path
from datetime import datetime, timedelta
try:
    from zoneinfo import ZoneInfo  # Python 3.9+
except Exception:
    ZoneInfo = None  # fallback para Python < 3.9
import re
from typing import Optional, Tuple, List, Dict, Any
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.options import Options as FirefoxOptions
from selenium.webdriver.support.ui import WebDriverWait as W
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException, WebDriverException, NoSuchElementException
)
from selenium.webdriver import ActionChains  # hover para abrir submenu
import re
import pandas as pd
from pathlib import Path
from datetime import datetime

# ---- Email (anexo) ----
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import formataddr
from email.mime.base import MIMEBase
from email import encoders

# ---------------------------------------------------------------------------------
# CONFIG
# ---------------------------------------------------------------------------------
BASE_URL_HOME = "https://www.iobonline.com.br/"
URL_CORE_HOME = "https://www.iobonline.com.br/pages/coreonline/integracao/issqn.jsf"
URL_MEUS_ALERTAS = "https://www.iobonline.com.br/pages/core/coreMeuEspacoAlertas.jsf"

WAIT_SEC = 20
HEADLESS = False

# Controles
RUN_CLICK_HISTORICO = True
HISTORICO_INDEX = 0  # fallback
RUN_CLICK_VER_DETALHES_DO_DIA = True

# Nome alvo do alerta a ter o sino clicado
ALERT_NAME_TARGET = "Alerta Padrão - ISSQN"

# Perfil real do Firefox
FIREFOX_PROFILE_PATH = r"C:\Users\a-81006408\AppData\Roaming\Mozilla\Firefox\Profiles\77teivsv.default"

DEFAULT_TZ = "America/Sao_Paulo"

# Saída em Excel
OUT_DIR = r"C:\Users\a-81006408\OneDrive - Vale S.A\Documentos\Dados Controles\IOB"
OUT_TEMP = Path(OUT_DIR, "Temp_base_atos_extraidos.xlsx")
OUT_BASE = Path(OUT_DIR, "Base_atos_extraidos.xlsx")
OUT_BACKUP = Path(OUT_DIR, "BACKUP_Base_atos_extraidos.xlsx")

# Metadados
FONTE_FIXA = "IOB - FEDERAL"         # fallback genérico
FONTE_MUNICIPAL = "IOB - MUNICIPAL"  # blocos municipais
ESFERA_FIXA = "FEDERAL"
ANALISTA_FIXO = None  # removido do layout
STATUS_CARGA_FIXO = "novo"  # pode trocar para "" se quiser

# Regex úteis
BR_RX = re.compile(r'(?:\s*<br\s*/?>\s*)+', re.IGNORECASE)  # <br> -> \n
MUNICIPAL_HDR_RX = re.compile(r'^\s*ISSQN\s*-\s*([A-Z]{2})\s*-\s*(.+?)\s*$', re.IGNORECASE)
DATE_TAIL_RX = re.compile(r'^\d{2}[./]\d{2}[./]\d{4}$')          # exatamente 10 chars dd.mm.yyyy / dd/mm/yyyy
DATE_ANY_RX = re.compile(r'(\d{2}[./]\d{2}[./]\d{4})')           # em qualquer posição

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------------
# UTIL
# ---------------------------------------------------------------------------------
def ensure_out_dir() -> None:
    Path(OUT_DIR).mkdir(parents=True, exist_ok=True)

def human_sleep(min_s: float = 0.8, max_s: float = 1.6) -> None:
    import random
    time.sleep(random.uniform(min_s, max_s))

def safe_click(driver, elem) -> None:
    """Scroll + JS click para evitar overlay/offsets em elementos dinâmicos."""
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", elem)
    driver.execute_script("arguments[0].click();", elem)

def _get_now_tz(tz_name: str) -> datetime:
    if ZoneInfo is not None:
        try:
            return datetime.now(ZoneInfo(tz_name))
        except Exception:
            pass
    return datetime.now()

def today_str_for_iob(
    tz_name: str = DEFAULT_TZ,
    test_date: Optional[str] = None,  # "YYYY-MM-DD"
    days_offset: int = 1
) -> str:
    """Retorna a data no formato 'Nov 5, 2025' (formato do site)."""
    if test_date is not None:
        test_date = test_date.strip()
        if not test_date:
            raise ValueError("test_date está vazio. Use 'YYYY-MM-DD' ou remova o parâmetro.")
        dt = datetime.strptime(test_date, "%Y-%m-%d")
    else:
        dt = _get_now_tz(tz_name)
    if days_offset:
        dt = dt + timedelta(days=days_offset)
    month_abbr = [None, "Jan", "Feb", "Mar", "Apr", "May", "Jun",
                  "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
    return f"{month_abbr[dt.month]} {dt.day}, {dt.year}"

def iob_english_date_to_iso(s: str) -> str:
    """Converte 'Nov 5, 2025' -> '2025-11-05' (ISO) (usado quando necessário)."""
    try:
        dt = datetime.strptime(s.strip(), "%b %d, %Y")
        return dt.strftime("%Y-%m-%d")
    except Exception:
        return ""

def data_extracao_like_old() -> str:
    """Mantém o mesmo formato 'YYYY-DD-MM' usado anteriormente."""
    hoje = datetime.today().strftime('%d.%m.%Y')
    return datetime.strptime(hoje, '%d.%m.%Y').strftime('%Y-%d-%m')

def _normalize_spaces(s: str) -> str:
    return " ".join((s or "").split()).strip()

def _clean_html_text(html_fragment: str) -> str:
    """
    Recebe um innerHTML simples (de uma <td>), troca <br> por \n,
    remove tags básicas e normaliza espaços/linhas.
    """
    if not html_fragment:
        return ""
    s = BR_RX.sub("\n", html_fragment)
    s = re.sub(r'</?(?:strong|u|b|i|em|span|font)[^>]*>', '', s, flags=re.IGNORECASE)
    s = re.sub(r'<[^>]+>', '', s)
    lines = [re.sub(r'\s+', ' ', ln).strip() for ln in s.splitlines()]
    lines = [ln for ln in lines if ln]
    return "\n".join(lines).strip()

def _try_parse_municipal_header(text: str):
    """Retorna (UF, Município) para 'ISSQN - UF - Município', senão None."""
    if not text:
        return None
    m = MUNICIPAL_HDR_RX.match(text.strip())
    if not m:
        return None
    uf = m.group(1).upper()
    municipio = m.group(2).strip()
    return uf, municipio

def extract_pub_date_from_ato_tail(ato: str) -> str:
    """
    Municipal: extrai data dos últimos 10 caracteres do Ato.
    Se não casar, tenta achar uma data dd.mm.yyyy ou dd/mm/yyyy nos últimos 40 chars.
    Retorna string (mantém o formato dd.mm.yyyy ou dd/mm/yyyy); "" se não encontrar.
    """
    if not ato:
        return ""
    tail = (ato[-10:]).strip()
    if DATE_TAIL_RX.match(tail):
        return tail
    # tenta achar próximo no final do título
    tail_window = ato[-40:]
    m = DATE_ANY_RX.search(tail_window)
    return m.group(1) if m else ""

# ---------------------------------------------------------------------------------
# DRIVER (perfil real)
# ---------------------------------------------------------------------------------
def build_driver_with_profile(headless: bool = HEADLESS):
    opts = FirefoxOptions()
    if headless:
        opts.add_argument("-headless")
    #opts.add_argument("-profile")
    #opts.add_argument(FIREFOX_PROFILE_PATH)
    try:
        opts.set_preference("dom.webdriver.enabled", False)
        opts.set_preference("useAutomationExtension", False)
        opts.set_preference("privacy.trackingprotection.enabled", False)
        opts.set_preference(
            "general.useragent.override",
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:121.0) Gecko/20100101 Firefox/121.0"
        )
    except Exception:
        pass
    driver = webdriver.Firefox(options=opts)
    driver.set_page_load_timeout(60)
    try:
        driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
    except Exception:
        pass
    logger.info("Firefox iniciado com perfil real.")
    return driver

# ---------------------------------------------------------------------------------
# LOGIN
# ---------------------------------------------------------------------------------
def accept_cookies_if_present(driver) -> None:
    try:
        btns = driver.find_elements(By.CSS_SELECTOR, "#onetrust-accept-btn-handler")
        for b in btns:
            if b.is_displayed():
                try:
                    b.click()
                    logger.info("Cookies aceitos (OneTrust).")
                except Exception:
                    pass
    except Exception:
        pass

def login_iob_simple(driver, user: str, pwd: str) -> bool:
    driver.get(BASE_URL_HOME)
    accept_cookies_if_present(driver)
    human_sleep()
    try:
        login_btn = W(driver, WAIT_SEC).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "a.button.button-login.font-button.enter"))
        )
        safe_click(driver, login_btn)
    except TimeoutException:
        logger.error("Botão 'Login' não encontrado.")
        return False

    human_sleep()
    try:
        email_input = W(driver, WAIT_SEC).until(
            EC.visibility_of_element_located((By.CSS_SELECTOR, "#txtLogin"))
        )
        email_input.clear()
        email_input.send_keys(user)
    except TimeoutException:
        logger.error("Campo #txtLogin não encontrado.")
        return False

    human_sleep()
    try:
        pwd_input = W(driver, WAIT_SEC).until(
            EC.visibility_of_element_located((By.CSS_SELECTOR, "#txtPassword"))
        )
        pwd_input.clear()
        pwd_input.send_keys(pwd)
    except TimeoutException:
        logger.error("Campo #txtPassword não encontrado.")
        return False

    human_sleep()
    try:
        entrar_btn = W(driver, WAIT_SEC).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "button.default-btn.light-green-btn.send-login"))
        )
        safe_click(driver, entrar_btn)
    except TimeoutException:
        logger.error("Botão 'Entrar' não encontrado.")
        return False

    human_sleep(2, 3.5)
    # Modal de sessão
    try:
        modal_btns = driver.find_elements(By.XPATH, "//span[contains(., 'Encerrar a sessão e logar')]")
        for b in modal_btns:
            if b.is_displayed():
                logger.info("Modal detectado — clicando em 'Encerrar a sessão e logar'.")
                safe_click(driver, b)
                human_sleep(1.5, 2.5)
                break
    except Exception:
        pass

    try:
        W(driver, WAIT_SEC * 2).until(
            EC.any_of(
                EC.url_contains("/area/"),
                EC.url_contains("/home"),
                EC.presence_of_element_located((By.XPATH, "//*[contains(translate(., 'SAIR','sair'),'sair')]"))
            )
        )
        logger.info("✅ Login realizado com sucesso.")
        return True
    except TimeoutException:
        logger.warning("⚠️ Login não confirmado (pode ter pedido CAPTCHA/manual).")
        return False

# ---------------------------------------------------------------------------------
# MENU: Meu Espaço -> Meus Alertas (via UI)
# ---------------------------------------------------------------------------------
def open_meu_espaco_and_click_meus_alertas(driver, wait_sec: int = WAIT_SEC) -> bool:
    try:
        xp_menu_label = "//div[contains(@class,'rich-label-text-decor')][contains(normalize-space(.),'Meu Espaço')]"
        menu_label = W(driver, wait_sec).until(
            EC.presence_of_element_located((By.XPATH, xp_menu_label))
        )
        try:
            ActionChains(driver).move_to_element(menu_label).pause(0.2).perform()
            human_sleep(0.2, 0.5)
        except Exception:
            pass
        try:
            W(driver, wait_sec).until(EC.element_to_be_clickable(menu_label))
        except Exception:
            pass
        safe_click(driver, menu_label)
        human_sleep(0.3, 0.8)

        xp_item = ("//span[contains(@class,'rich-menu-item-label') and "
                   "(normalize-space(.)='Meus Alertas' or contains(@id,':meusAlertas:anchor'))]")
        item_span = W(driver, wait_sec).until(
            EC.visibility_of_element_located((By.XPATH, xp_item))
        )
        try:
            link = item_span.find_element(By.XPATH, "ancestor::a[1]")
        except NoSuchElementException:
            link = item_span
        try:
            W(driver, wait_sec).until(EC.element_to_be_clickable(link))
        except Exception:
            pass
        safe_click(driver, link)
        human_sleep(0.6, 1.2)
        try:
            W(driver, wait_sec).until(
                EC.any_of(
                    EC.url_contains("coreMeuEspacoAlertas.jsf"),
                    EC.presence_of_element_located((By.XPATH, "//*[contains(., 'MEUS ALERTAS') or contains(., 'Resultados:')]"))
                )
            )
        except TimeoutException:
            logger.warning("Não consegui confirmar o carregamento de 'Meus Alertas' (pode estar oculto no HTML).")
        logger.info("✅ Menu 'Meu Espaço' -> 'Meus Alertas' aberto com sucesso.")
        return True
    except TimeoutException:
        logger.error("Não localizei 'Meu Espaço' ou 'Meus Alertas' dentro do tempo.")
        return False
    except Exception as e:
        logger.exception(f"Falha ao abrir 'Meus Alertas' via menu: {e}")
        return False

# ---------------------------------------------------------------------------------
# HISTÓRICO (sino)
# ---------------------------------------------------------------------------------
def click_historico(driver, index: int = 0, wait_sec: int = WAIT_SEC) -> bool:
    """Fallback por índice (primeiro ícone 'Histórico' da página)."""
    try:
        W(driver, wait_sec).until(
            EC.presence_of_element_located(
                (By.XPATH, "//img[@alt='Histórico' and contains(@src,'ico_alerta.gif')]")
            )
        )
        elems = driver.find_elements(By.XPATH, "//img[@alt='Histórico' and contains(@src,'ico_alerta.gif')]")
        if not elems:
            logger.warning("Nenhum ícone 'Histórico' encontrado.")
            return False
        if index < 0 or index >= len(elems):
            logger.warning("Índice fora do intervalo.")
            return False
        img = elems[index]
        try:
            link = img.find_element(By.XPATH, "ancestor::a[1]")
        except NoSuchElementException:
            link = img
        try:
            W(driver, wait_sec).until(EC.element_to_be_clickable(link))
        except Exception:
            pass
        safe_click(driver, link)
        human_sleep(0.6, 1.0)
        logger.info("Clique no ícone 'Histórico' (por índice) realizado.")
        try:
            W(driver, 10).until(
                EC.any_of(
                    EC.url_contains("AlertasHistory"),
                    EC.presence_of_element_located((By.XPATH, "//table//tbody//tr")),
                    EC.staleness_of(img)
                )
            )
        except TimeoutException:
            pass
        return True
    except TimeoutException:
        logger.warning("Tempo esgotado esperando o ícone 'Histórico'.")
        return False
    except Exception as e:
        logger.exception(f"Falha ao clicar no 'Histórico': {e}")
        return False

def click_historico_by_alert_name(driver, alert_name: str, wait_sec: int = WAIT_SEC) -> bool:
    """
    Clica o sino (ícone 'Histórico') na TR cujo 'Nome do Alerta' contém alert_name.
    """
    target = _normalize_spaces(alert_name).lower()
    logger.info("Procurando sino do alerta com nome contendo: %r", alert_name)
    try:
        try:
            W(driver, wait_sec).until(
                EC.presence_of_element_located((
                    By.XPATH,
                    "//table[.//th[contains(normalize-space(.),'Nome do Alerta')] or "
                    ".//td[contains(normalize-space(.),'Nome do Alerta')]]"
                ))
            )
        except TimeoutException:
            logger.warning("Não localizei a tabela 'Nome do Alerta'. Tentando procurar linhas mesmo assim.")
        rows = driver.find_elements(By.XPATH, "//table//tr[.//td]")
        if not rows:
            logger.warning("Nenhuma linha de alerta foi encontrada na página.")
            return False
        for row in rows:
            try:
                name_cells = row.find_elements(
                    By.XPATH, ".//td[contains(@class,'textoDefaultProduto') or .//text()]"
                )
                row_name_text = " ".join([_normalize_spaces(c.text) for c in name_cells]).lower()
                if not row_name_text:
                    row_name_text = _normalize_spaces(row.text).lower()
                if target and target in row_name_text:
                    try:
                        img = row.find_element(
                            By.XPATH,
                            ".//img[@alt='Histórico' and contains(@src,'ico_alerta.gif')]"
                        )
                    except NoSuchElementException:
                        logger.warning("Linha encontrada, mas não achei o ícone 'Histórico'.")
                        return False
                    try:
                        link = img.find_element(By.XPATH, "ancestor::a[1]")
                    except NoSuchElementException:
                        link = img
                    try:
                        W(driver, wait_sec).until(EC.element_to_be_clickable(link))
                    except Exception:
                        pass
                    safe_click(driver, link)
                    human_sleep(0.4, 0.9)
                    try:
                        W(driver, 10).until(
                            EC.any_of(
                                EC.url_contains("AlertasHistory"),
                                EC.presence_of_element_located((By.XPATH, "//*[contains(., 'Resultados:')]")),
                                EC.staleness_of(img)
                            )
                        )
                    except TimeoutException:
                        pass
                    logger.info("✅ Sino do alerta clicado com sucesso (por nome).")
                    return True
            except Exception:
                continue
        logger.warning("Não encontrei linha cujo nome contenha: %r", alert_name)
        return False
    except Exception as e:
        logger.exception("Erro ao clicar no sino por nome do alerta: %s", e)
        return False

# ---------------------------------------------------------------------------------
# VER DETALHES (por data)
# ---------------------------------------------------------------------------------
def click_ver_detalhes_for_date(driver, date_str: str, wait_sec: int = WAIT_SEC) -> bool:
    logger.info("Procurando linha com data: %r", date_str)
    try:
        try:
            W(driver, wait_sec).until(
                EC.presence_of_element_located((By.XPATH, "//*[contains(., 'MEUS ALERTAS') or contains(., 'Resultados:')]"))
            )
        except TimeoutException:
            pass
        xpath_row = "//tr[.//div[@align='center' and normalize-space(.)='%s']]" % date_str
        row = W(driver, wait_sec).until(
            EC.presence_of_element_located((By.XPATH, xpath_row))
        )
        try:
            btn = row.find_element(By.XPATH, ".//img[contains(@src,'bt_ver_detalhes.gif')]")
        except NoSuchElementException:
            try:
                btn = row.find_element(
                    By.XPATH,
                    ".//*[self::img or self::input][contains(@src,'bt_ver_detalhes')] | "
                    ".//a[contains(normalize-space(.), 'Ver detalhes')]"
                )
            except NoSuchElementException:
                logger.warning("Botão 'Ver detalhes' não encontrado na linha da data.")
                return False
        try:
            link = btn.find_element(By.XPATH, "ancestor::a[1]")
        except NoSuchElementException:
            link = btn
        try:
            W(driver, wait_sec).until(EC.element_to_be_clickable(link))
        except Exception:
            pass
        safe_click(driver, link)
        human_sleep(0.6, 1.1)
        logger.info("Clique em 'Ver detalhes' da data %s realizado.", date_str)
        return True
    except TimeoutException:
        logger.info("Não encontrei linha com a data exata %s.", date_str)
        return False
    except Exception as e:
        logger.exception("Erro ao clicar em 'Ver detalhes' para %s: %s", date_str, e)
        return False

def click_ver_detalhes_for_today(driver,
                                 tz_name: str = DEFAULT_TZ,
                                 test_date: Optional[str] = None,
                                 days_offset: int = 0,
                                 wait_sec: int = WAIT_SEC) -> Tuple[bool, str]:
    date_str = today_str_for_iob(tz_name=tz_name, test_date=test_date, days_offset=days_offset)
    logger.info("🔎 Procurando pela data: %r", date_str)
    ok = click_ver_detalhes_for_date(driver, date_str, wait_sec=wait_sec)
    return ok, date_str

# ---------------------------------------------------------------------------------
# EXTRAÇÃO (detalhes do dia) — Parser municipal + fallback
# ---------------------------------------------------------------------------------
def gather_item_cards_in_details(driver):
    """Fallback genérico: tenta linhas de tabela, articles e listas."""
    cards = []
    rows = driver.find_elements(By.XPATH, "//table//tbody//tr[.//a or .//h3 or .//p]")
    if rows:
        cards.extend(rows)
    if not cards:
        arts = driver.find_elements(By.XPATH, "//article[contains(@class,'principaisnormas')]//article")
        if arts:
            cards.extend(arts)
    if not cards:
        lis = driver.find_elements(By.XPATH, "//ul/li[.//a or .//h3 or .//p]")
        if lis:
            cards.extend(lis)
    return cards

def extract_title_and_snippet_from_card(card):
    """Fallback genérico: extrai título/snippet/link com seletores tolerantes."""
    title = None
    body = None
    for xp in [".//h3", ".//h2", ".//h1",
               ".//a[contains(@class,'titulo') or contains(@class,'title') or contains(@class,'link')][1]",
               ".//a[1]"]:
        try:
            el = card.find_element(By.XPATH, xp)
            title = el.text.strip()
            if title:
                break
        except Exception:
            pass
    for xp in [".//p[1]", ".//div[contains(@class,'resumo') or contains(@class,'snippet')][1]"]:
        try:
            el = card.find_element(By.XPATH, xp)
            body = el.text.strip()
            if body:
                break
        except Exception:
            pass
    link = None
    for xp in [".//a[contains(@href,'http')][1]", ".//a[1]"]:
        try:
            el = card.find_element(By.XPATH, xp)
            href = el.get_attribute("href") or ""
            if href:
                link = href
                break
        except Exception:
            pass
    return title, body, link

def extract_items_municipal_blocks(driver, target_date_str: str) -> List[Dict[str, Any]]:
    """
    Parser específico para blocos municipais:
    <strong>ISSQN - UF - Município</strong>
    <strong>Decreto nº ...</strong>
    <td class="txt_cinza"> ... <br><br>DESCRIÇÃO...</td>
    """
    items: List[Dict[str, Any]] = []
    strong_nodes = driver.find_elements(By.XPATH, "//strong")
    if not strong_nodes:
        return items

    current_uf: Optional[str] = None
    current_municipio: Optional[str] = None

    blacklist_header_snippets = ["Caro(a)", "Seguem abaixo as atualizações", "Título do alerta"]

    for s in strong_nodes:
        try:
            txt = (s.text or "").strip()
            if not txt:
                continue

            # Cabeçalho 'ISSQN - UF - Município'
            hdr = _try_parse_municipal_header(txt)
            if hdr:
                current_uf, current_municipio = hdr
                continue

            # Se há UF/município ativo, este <strong> é o ATO do item
            if current_uf and current_municipio:
                ato = txt

                # DESCRIÇÃO: preferir o <td class="txt_cinza"> ancestral do próprio <strong>
                descricao = ""
                desc_node = None

                # 1) ancestral <td.txt_cinza>
                try:
                    td_anc = s.find_element(By.XPATH, "ancestor::td[contains(@class,'txt_cinza')][1]")
                    desc_node = td_anc
                except Exception:
                    desc_node = None

                # 2) mesma linha
                if desc_node is None:
                    try:
                        desc_node = s.find_element(
                            By.XPATH,
                            "ancestor::tr[1]//td[contains(@class,'txt_cinza')][1]"
                        )
                    except Exception:
                        desc_node = None

                # 3) próxima linha
                if desc_node is None:
                    try:
                        desc_node = s.find_element(
                            By.XPATH,
                            "ancestor::tr[1]/following-sibling::tr[1]//td[contains(@class,'txt_cinza')][1]"
                        )
                    except Exception:
                        desc_node = None

                # 4) fallback: próximo td.txt_cinza (com filtro contra cabeçalhos)
                if desc_node is None:
                    try:
                        candidate = s.find_element(By.XPATH, "following::td[contains(@class,'txt_cinza')][1]")
                        cand_text = (candidate.text or "").strip()
                        if not any(key in cand_text for key in blacklist_header_snippets):
                            desc_node = candidate
                    except Exception:
                        desc_node = None

                if desc_node is not None:
                    try:
                        inner = desc_node.get_attribute("innerHTML") or desc_node.text
                    except Exception:
                        inner = desc_node.text

                    desc_full = _clean_html_text(inner)
                    # Remove a linha do próprio ATO, mantendo apenas os parágrafos
                    lines = [ln.strip() for ln in desc_full.splitlines() if ln.strip()]
                    norm_ato = _normalize_spaces(ato).lower()
                    lines = [ln for ln in lines if _normalize_spaces(ln).lower() != norm_ato]
                    descricao = "\n".join(lines).strip()

                # Data de publicação: últimos 10 chars do Ato (heurística municipal)
                data_publicacao = extract_pub_date_from_ato_tail(ato)

                # Data de extração
                data_extracao = data_extracao_like_old()

                item = {
                    "Ato": ato,
                    "Descrição": (descricao or "").strip(),
                    "Esfera": "MUNICIPAL",
                    "UF": current_uf,
                    "Municipio": current_municipio,
                    "Data de extração": data_extracao,
                    "Data de publicação": data_publicacao,
                    "Fonte": FONTE_MUNICIPAL,
                    "StatusCarga": STATUS_CARGA_FIXO,
                }
                items.append(item)
        except Exception:
            continue

    return items

def extract_items_from_details_page(driver, target_date_str: str) -> List[Dict[str, Any]]:
    """
    Extrai itens com duas estratégias:
    1) Parser de blocos municipais (prioritário).
    2) Fallback genérico (tabelas/articles/listas).
    """
    # 1) TENTATIVA: blocos municipais
    try:
        items_blocks = extract_items_municipal_blocks(driver, target_date_str=target_date_str)
    except Exception as e:
        logger.debug("Parser municipal falhou: %s", e)
        items_blocks = []
    if items_blocks:
        logger.info("Extração (blocos municipais): %d itens.", len(items_blocks))
        return items_blocks

    # 2) FALLBACK: estratégia genérica
    logger.info("Nenhum bloco municipal detectado. Usando fallback genérico.")
    items: List[Dict[str, Any]] = []
    cards = gather_item_cards_in_details(driver)
    if not cards:
        logger.warning("Nenhum item encontrado nos detalhes (fallback).")
        return items

    for card in cards:
        try:
            title, snippet, link = extract_title_and_snippet_from_card(card)
            if not title and not snippet:
                continue

            full_title, full_body = (None, None)
            if link:
                try:
                    full_title, full_body = open_link_and_extract_full_text(driver, link)
                    human_sleep(0.4, 0.9)
                except Exception as e:
                    logger.debug("Falha ao abrir link do item: %s", e)

            ato = (full_title or title or "").strip()
            descricao = (full_body or snippet or "").strip()

            # Data de publicação (melhor esforço): tenta extrair do título também
            data_publicacao = extract_pub_date_from_ato_tail(ato)

            item = {
                "Ato": ato,
                "Descrição": descricao,
                "Esfera": ESFERA_FIXA,     # FEDERAL
                "UF": "FEDERAL",
                "Municipio": "",
                "Data de extração": data_extracao_like_old(),
                "Data de publicação": data_publicacao,  # pode ficar "" se não houver
                "Fonte": FONTE_FIXA,
                "StatusCarga": STATUS_CARGA_FIXO,
            }
            items.append(item)
        except Exception as e:
            logger.debug("Falha ao processar um card (fallback): %s", e)

    logger.info("Extração (fallback): %d itens.", len(items))
    return items

# ---------------------------------------------------------------------------------
# EXCEL (novo layout)
# ---------------------------------------------------------------------------------
FINAL_COL_ORDER = [
    "Ato",
    "Descrição",
    "Esfera",
    "UF",
    "Municipio",
    "Data de extração",
    "Data de publicação",
    "Fonte",
    "StatusCarga",
]

# ---- COLUNAS PARA DEDUPLICAÇÃO (ignorando "Data de extração") ----
DEDUP_COLS = [
    "Ato",
    "Descrição",
    "Esfera",
    "UF",
    "Municipio",
    "Data de publicação",
    "Fonte",
]

# ---- HELPERS ----
def _norm_text(s: str) -> str:
    """Normaliza texto para comparação: trim + colapsa espaços internos."""
    if pd.isna(s):
        return ""
    s = str(s)
    s = re.sub(r"\s+", " ", s.strip())
    return s

def dedupe_base_excel(input_path: Path, output_path: Path) -> int:
    """
    Lê a base consolidada, normaliza as colunas de comparação e remove duplicados por DEDUP_COLS.
    Salva a planilha resultante no output_path. Retorna a quantidade removida.
    """
    if not input_path.exists():
        logger.warning("Arquivo base não encontrado para dedupe: %s", input_path)
        return 0

    df = pd.read_excel(input_path, engine="openpyxl")

    # Garante presença das colunas-chave
    for col in DEDUP_COLS:
        if col not in df.columns:
            df[col] = ""

    # Normaliza apenas para comparação
    df_cmp = df.copy()
    for col in DEDUP_COLS:
        df_cmp[col] = df_cmp[col].apply(_norm_text)

    before = len(df_cmp)
    df_dedup = df_cmp.drop_duplicates(subset=DEDUP_COLS, keep="first")
    removed = before - len(df_dedup)

    # Mantém a ordem final de colunas, criando as ausentes se preciso
    for col in FINAL_COL_ORDER:
        if col not in df_dedup.columns:
            df_dedup[col] = ""
    df_out = df_dedup[FINAL_COL_ORDER]

    output_path.parent.mkdir(parents=True, exist_ok=True)
    df_out.to_excel(output_path, sheet_name="dados", index=False)
    logger.info("Base deduplicada salva em: %s (removidos: %d)", output_path, removed)
    return removed

def send_mail_with_attachment(
    smtp_server: str,
    smtp_port: int,
    from_addr: str,
    from_name: str,
    to_addr: str,
    bcc_addr: str,
    subject: str,
    body: str,
    attachment_path: Path
) -> None:
    """Envia e-mail sem SSL (porta 25) com arquivo em anexo."""
    msg = MIMEMultipart()
    msg["From"] = formataddr((from_name, from_addr))
    msg["To"] = to_addr
    msg["Subject"] = subject
    msg.attach(MIMEText(body, "plain"))

    # Anexo (se existir)
    if attachment_path and attachment_path.exists():
        with open(attachment_path, "rb") as f:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(f.read())
        encoders.encode_base64(part)
        part.add_header(
            "Content-Disposition",
            f'attachment; filename="{attachment_path.name}"'
        )
        msg.attach(part)
    else:
        logger.warning("Anexo não encontrado: %s", attachment_path)

    # Envio
    try:
        with smtplib.SMTP(smtp_server, smtp_port, timeout=30) as server:
            server.send_message(msg)
        logger.info("E-mail enviado com sucesso para %s", to_addr)
    except Exception as e:
        logger.exception("Falha ao enviar e-mail: %s", e)

# ---- FUNÇÃO PRINCIPAL DE SALVAMENTO + PÓS-PROCESSAMENTO ----
def save_to_excel_like_old(items: List[Dict[str, Any]]) -> None:
    """
    Cria DataFrame no novo layout, salva Temp_Orientacoes_iob.xlsx e consolida em Base_Orientacoes.xlsx (com backup),
    removendo duplicados por ['Ato','Descrição','Fonte'] na consolidação.
    Em seguida, executa uma deduplicação final ignorando 'Data de extração' (colunas DEDUP_COLS),
    salva como 'Base_atos_extraidos.xlsx' e envia por e-mail com o anexo.
    """
    if not items:
        logger.info("Nenhum item para salvar.")
        return

    ensure_out_dir()
    df = pd.DataFrame(items)

    # Garante ordem e presença das colunas finais
    for col in FINAL_COL_ORDER:
        if col not in df.columns:
            df[col] = ""
    df = df[FINAL_COL_ORDER]

    # 1) arquivo temporário
    try:
        df.to_excel(OUT_TEMP, sheet_name="dados", index=False)
        logger.info("Temp salvo: %s", OUT_TEMP)
    except Exception as e:
        logger.error("Falha ao salvar temp: %s", e)

    # 2) consolidar base (append + drop_duplicates simples)
    try:
        if OUT_BASE.exists():
            df_base = pd.read_excel(OUT_BASE, engine="openpyxl")
            # garante mesmas colunas/ordem na base
            for col in FINAL_COL_ORDER:
                if col not in df_base.columns:
                    df_base[col] = ""
            df_base = df_base[FINAL_COL_ORDER]
            df_all = pd.concat([df_base, df], ignore_index=True)
        else:
            df_all = df.copy()

        # Dedupe primário (como você já fazia)
        df_all = df_all.drop_duplicates(subset=["Ato", "Descrição", "Fonte"], keep="first")
        df_all.to_excel(OUT_BASE, sheet_name="dados", index=False)
        logger.info("Base consolidada: %s", OUT_BASE)

        # backup
        try:
            df_all.to_excel(OUT_BACKUP, sheet_name="dados", index=False)
            logger.info("Backup salvo: %s", OUT_BACKUP)
        except Exception as e:
            logger.warning("Falha ao salvar backup: %s", e)
    except Exception as e:
        logger.error("Falha na consolidação da base: %s", e)
        # Se falhar aqui, não segue para o pós-processamento
        return

    # 3) PÓS-PROCESSAMENTO: deduplicar ignorando "Data de extração" e enviar e-mail
    try:
        # Caminho da base consolidada que acabamos de gerar
        base_consolidada = OUT_BASE

        # Saída deduplicada com o nome solicitado
        base_deduplicada = Path(OUT_DIR, "Base_atos_extraidos.xlsx")

        removidos = dedupe_base_excel(
            input_path=base_consolidada,
            output_path=base_deduplicada
        )

        # Monta assunto e corpo do e-mail
        hoje = datetime.now().strftime("%Y-%m-%d")
        subject = f"[IOB] Base extraída ({hoje}) - Duplicados removidos: {removidos}"
        body = (
            "Olá,\n\n"
            f"A base de atos foi gerada e deduplicada com sucesso em {hoje}.\n"
            f"Registros duplicados removidos (ignorando 'Data de extração'): {removidos}.\n\n"
            "Segue a planilha em anexo.\n\n"
            "Att.,\nRobô IOB"
        )

        # Parâmetros de e-mail (conforme os seus)
        SMTP_SERVER = "mailbr.valeglobal.net"
        SMTP_PORT = 25  # SMTP sem SSL
        FROM_ADDR = "servicos.contabeis@vale.com"
        FROM_NAME = "Serviços Contábeis (Robô IOB)"
        TO_ADDR = "diego.r.lemos@vale.com"
        #"bianca.castro@vale.com, carolina.romanelli.souza@vale.com, joao.los@vale.com"
        BCC_ADDR = "diego.r.lemos@vale.com"
        send_mail_with_attachment(
            smtp_server=SMTP_SERVER,
            smtp_port=SMTP_PORT,
            from_addr=FROM_ADDR,
            from_name=FROM_NAME,
            to_addr=TO_ADDR,
            bcc_addr=BCC_ADDR,
            subject=subject,
            body=body,
            attachment_path=base_deduplicada
        )
    except Exception as e:
                logger.exception("Falha no pós-processamento (dedupe/email): %s", e)
# ---------------------------------------------------------------------------------
# AUX: abrir link do item e extrair título/corpo
# ---------------------------------------------------------------------------------
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

def open_link_and_extract_full_text(driver, url: str, wait_sec: int = WAIT_SEC) -> Tuple[Optional[str], Optional[str]]:
    """
    Abre o link em nova aba, raspa título (h1/h2) e corpo (parágrafos).
    Fecha a aba e retorna ao contexto original.
    """
    original = driver.current_window_handle
    driver.execute_script("window.open(arguments[0], '_blank');", url)
    W(driver, wait_sec).until(EC.number_of_windows_to_be(2))
    new_tab = [h for h in driver.window_handles if h != original][0]
    driver.switch_to.window(new_tab)
    try:
        W(driver, wait_sec).until(
            EC.presence_of_all_elements_located((By.XPATH, "(//h1|//h2|//article|//div)"))
        )
        full_title = None
        for xp in ["//article//h1", "//h1", "//article//h2", "//h2"]:
            try:
                el = driver.find_element(By.XPATH, xp)
                full_title = el.text.strip()
                if full_title:
                    break
            except Exception:
                pass

        paragraphs: List[str] = []
        for xp in ["//article//p",
                   "//div[contains(@class,'conteudo') or contains(@class,'content') or contains(@class,'texto')]//p",
                   "//section//p", "//p"]:
            try:
                els = driver.find_elements(By.XPATH, xp)
                for p in els:
                    txt = p.text.strip()
                    if txt:
                        paragraphs.append(txt)
                if paragraphs:
                    break
            except Exception:
                pass
        full_body = "\n\n".join(paragraphs) if paragraphs else None
        return full_title, full_body
    finally:
        try:
            driver.close()
        except Exception:
            pass
        driver.switch_to.window(original)

# ---------------------------------------------------------------------------------
# MAIN
# ---------------------------------------------------------------------------------
def main() -> None:
    # credenciais via arquivo .ENV ou variáveis de ambiente
    env = load_env_if_exists(r"C:\Users\a-81006408\PycharmProjects\IOBP.ENV") or {}
    user = env.get("IOB_EMAIL") or os.environ.get("IOB_EMAIL")
    pwd = env.get("IOB_SENHA") or os.environ.get("IOB_SENHA")
    if not user:
        user = input("IOB_EMAIL: ").strip()
    if not pwd:
        pwd = input("IOB_SENHA: ").strip()

    driver = None
    try:
        driver = build_driver_with_profile(HEADLESS)
        success = login_iob_simple(driver, user, pwd)
        if not success:
            logger.info("Falha no login. Verifique CAPTCHA/seletores.")
            time.sleep(4)
            return

        # Core Home
        logger.info("Logado — abrindo Core Home.")
        driver.get(URL_CORE_HOME)
        try:
            W(driver, WAIT_SEC).until(EC.url_contains("coreHome.jsf"))
        except TimeoutException:
            logger.warning("Não foi possível confirmar o carregamento da Core Home.")
        human_sleep(0.4, 0.9)

        # Meu Espaço -> Meus Alertas (via menu)
        logger.info("Abrindo Meu Espaço > Meus Alertas pelo menu.")
        ok_menu = open_meu_espaco_and_click_meus_alertas(driver, wait_sec=WAIT_SEC)
        if not ok_menu:
            logger.error("Não consegui abrir 'Meus Alertas' via menu. Encerrando.")
            return

        # Clicar o sino do alerta-alvo
        if RUN_CLICK_HISTORICO and ALERT_NAME_TARGET:
            ok_hist = click_historico_by_alert_name(driver, alert_name=ALERT_NAME_TARGET, wait_sec=WAIT_SEC)
            if not ok_hist:
                logger.warning("Não consegui clicar o sino pelo nome. Tentando fallback por índice...")
                ok_hist = click_historico(driver, index=HISTORICO_INDEX, wait_sec=WAIT_SEC)
        try:
            W(driver, 10).until(
                EC.any_of(
                    EC.url_contains("AlertasHistory"),
                    EC.presence_of_element_located((By.XPATH, "//table//tbody//tr"))
                )
            )
        except TimeoutException:
            pass

        # Ver detalhes da data
        FORCE_TEST_DATE: Optional[str] = None
        DAYS_OFFSET = 0  # ajuste conforme necessário (0=hoje)
        ok_dia, target_date_str = click_ver_detalhes_for_today(
            driver, tz_name=DEFAULT_TZ, test_date=FORCE_TEST_DATE, days_offset=DAYS_OFFSET
        )
        if not ok_dia:
            logger.warning("Não há linha para a data alvo OU o botão 'Ver detalhes' não foi encontrado.")
            time.sleep(3)
            return

        # EXTRAÇÃO (prioriza parser municipal; fallback se preciso)
        items = extract_items_from_details_page(driver, target_date_str=target_date_str)
        logger.info("Itens extraídos: %d", len(items))

        # SALVAR (novo layout)
        save_to_excel_like_old(items)
        time.sleep(2)

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