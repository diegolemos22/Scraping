
# -*- coding: utf-8 -*-
"""
INSERTs em SharePoint Online (REST + cookies, sem OAuth):
- Abre Firefox com perfil existente e extrai __REQUESTDIGEST do DOM (página clássica)
- Exporta cookies do Firefox para requests, reforçando FedAuth em path='/'
- Resolve lista por caminho, mapeia Title->InternalName e cria itens via POST (odata=nometadata, sem __metadata)
"""

import os
import time
import re
import html
from urllib.parse import quote

import pandas as pd
import requests
from selenium import webdriver
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.firefox.service import Service

# ==========================
# CONFIGURAÇÃO
# ==========================
# Seu perfil REAL do Firefox
FIREFOX_PROFILE_PATH = r"C:\Users\a-81006408\AppData\Roaming\Mozilla\Firefox\Profiles\77teivsv.default"

# (Opcional) Caminho do geckodriver, se não estiver no PATH:
GECKODRIVER_PATH = None  # ex.: r"C:\Tools\geckodriver.exe"

# Site e lista
SITE_URL = "https://globalvale.sharepoint.com/sites/SICOM-SistemaIntegradodeComodato"
LIST_SERVER_RELATIVE_URL = "/sites/SICOM-SistemaIntegradodeComodato/Lists/Requisies de Comodato"

# Página clássica para raspar o digest:
DIGEST_PAGE = SITE_URL + "/_layouts/15/settings.aspx"

# Excel
EXCEL_CAMINHO = r"C:\Users\a-81006408\PycharmProjects\Sql\EXPORT_SSIS\APP_RODIZIO\BASE-RODIZIO 3.xlsx"
PLANILHA_PRIORITARIA = None  # None -> primeira aba; ou "BASE"

# ==========================
# Iniciar Firefox com perfil existente (sem copiar)
# ==========================
def start_firefox_with_profile(profile_path):
    if not os.path.isdir(profile_path):
        raise FileNotFoundError(f"Perfil não encontrado: {profile_path}")

    opts = Options()
    # Usar o perfil real diretamente
    opts.add_argument("-no-remote")
    opts.add_argument("-profile")
    opts.add_argument(profile_path)

    if GECKODRIVER_PATH and os.path.isfile(GECKODRIVER_PATH):
        service = Service(GECKODRIVER_PATH)
        driver = webdriver.Firefox(service=service, options=opts)
    else:
        driver = webdriver.Firefox(options=opts)  # geckodriver deve estar no PATH

    driver.maximize_window()
    return driver

# ==========================
# Extrair __REQUESTDIGEST do DOM
# ==========================
def get_requestdigest_from_dom(driver, url):
    driver.get(url)
    time.sleep(5)  # aguarda carregar/autenticar

    # Tenta via JS:
    digest = driver.execute_script(
        "return document.getElementById('__REQUESTDIGEST')?.value || "
        "document.querySelector('input[name=__REQUESTDIGEST]')?.value || null;"
    )
    # Fallback: regex no HTML bruto
    if not digest:
        html_text = driver.page_source
        m = re.search(r'name="__REQUESTDIGEST"\s+value="([^"]+)"', html_text, flags=re.IGNORECASE)
        if m:
            digest = html.unescape(m.group(1))
    return digest

# ==========================
# Converter cookies do Firefox -> requests
# ==========================
def make_session_from_driver(driver):
    sess = requests.Session()
    sess.headers.update({
        "Accept": "application/json;odata=nometadata",
        "Content-Type": "application/json;odata=nometadata",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) PythonRequests",
        "Prefer": "return=representation",
    })

    names = []
    for c in driver.get_cookies():
        domain = c.get("domain") or "globalvale.sharepoint.com"
        path = c.get("path") or "/"
        sess.cookies.set(c["name"], c["value"], domain=domain, path=path)
        names.append(c["name"])

    # Reforça FedAuth/rtFa com path='/'
    def reforce(name):
        for c in driver.get_cookies():
            if c["name"] == name:
                sess.cookies.set(name, c["value"], domain="globalvale.sharepoint.com", path="/")
                return True
        return False

    fed_ok = reforce("FedAuth")
    rtfa_ok = reforce("rtFa")
    print("Cookies do Firefox:", sorted(set(names)))
    print("Reforçou FedAuth=/ ?", fed_ok, "| rtFa=/ ?", rtfa_ok)

    active = sess.cookies.get_dict(domain="globalvale.sharepoint.com")
    print("Cookies ativos no host:", list(active.keys()))
    print("Tem FedAuth para / ?", "FedAuth" in active)
    print("Tem rtFa   para / ?", "rtFa"   in active)

    return sess

# ==========================
# REST helpers (GET/POST)
# ==========================
def resolve_list_by_path(sess, site_url, server_relative_url):
    url_param = quote(server_relative_url.rstrip('/'), safe="/")
    getlist_url = f"{site_url}/_api/web/GetList(@listUrl)?@listUrl='{url_param}'"
    print("GetList URL:", getlist_url)
    r = sess.get(getlist_url)
    print("GetList status:", r.status_code, "| preview:", r.text[:180], "...")
    r.raise_for_status()
    info = r.json()  # nometadata -> sem 'd'
    return {"Id": info["Id"], "Title": info["Title"], "Path": server_relative_url.rstrip('/')}

def get_fields_title_to_internal(sess, site_url, list_guid):
    url = f"{site_url}/_api/web/lists(guid'{list_guid}')/fields?$select=Title,InternalName,Hidden,ReadOnlyField"
    r = sess.get(url)
    print("Fields status:", r.status_code, "| preview:", r.text[:180], "...")
    r.raise_for_status()
    data = r.json()  # nometadata -> 'value'
    mapping = {}
    for f in data.get("value", []):
        if not f.get("Hidden", False) and not f.get("ReadOnlyField", False):
            title = (f.get("Title") or "").strip()
            internal = f.get("InternalName")
            if title and internal:
                mapping[title] = internal
    print("Mapeamento Title->Internal:", mapping)
    return mapping

def create_item(sess, site_url, list_guid, payload_fields):
    url = f"{site_url}/_api/web/lists(guid'{list_guid}')/items"
    print("POST URL:", url, "| keys:", list(payload_fields.keys()))
    r = sess.post(url, json=payload_fields)
    print("POST status:", r.status_code, "| preview:", r.text[:220], "...")
    if r.status_code in (200, 201):
        try:
            return True, r.json()
        except Exception:
            return True, {"status": r.status_code}
    else:
        try:
            return False, r.json()
        except Exception:
            return False, r.text

# ==========================
# MAIN (INSERTS)
# ==========================
def main():
    # 1) Firefox com perfil logado (feche o Firefox antes de rodar)
    driver = start_firefox_with_profile(FIREFOX_PROFILE_PATH)

    # 2) Raspar __REQUESTDIGEST
    digest = get_requestdigest_from_dom(driver, DIGEST_PAGE)
    if not digest:
        # tenta a página da lista (alguns tenants exibem digest ali)
        digest = get_requestdigest_from_dom(driver, SITE_URL + "/Lists/Requisies%20de%20Comodato/AllItems.aspx")

    if not digest:
        driver.quit()
        raise RuntimeError("Não foi possível obter __REQUESTDIGEST no DOM. Abra a página clássica logado e tente novamente.")
    print("Digest (len):", len(digest))

    # 3) Exportar cookies do Firefox para requests
    sess = make_session_from_driver(driver)
    driver.quit()  # navegador não é mais necessário

    # 4) Anexar digest aos headers
    sess.headers.update({"X-RequestDigest": digest})

    # 5) Resolver lista por caminho (GUID + Title)
    info = resolve_list_by_path(sess, SITE_URL, LIST_SERVER_RELATIVE_URL)
    list_guid = info["Id"]
    print(f"Lista: '{info['Title']}' | GUID={list_guid}")

    # 6) Obter mapeamento Title->Internal
    title_to_internal = get_fields_title_to_internal(sess, SITE_URL, list_guid)

    # 7) Ler Excel
    if PLANILHA_PRIORITARIA:
        df = pd.read_excel(EXCEL_CAMINHO, sheet_name=PLANILHA_PRIORITARIA,
                           dtype=str, na_filter=False, engine='openpyxl')
        print(f"Lida planilha: {PLANILHA_PRIORITARIA}")
    else:
        xl = pd.ExcelFile(EXCEL_CAMINHO, engine='openpyxl')
        df = pd.read_excel(xl, sheet_name=xl.sheet_names[0],
                           dtype=str, na_filter=False, engine='openpyxl')
        print(f"Lida planilha: {xl.sheet_names[0]}")

    df.columns = df.columns.str.strip()
    df = df.fillna('')

    # 8) Filtrar por Titles conhecidos e renomear para InternalName
    common_titles = [t for t in df.columns if t in title_to_internal]
    if not common_titles:
        print("⚠ Nenhuma coluna do Excel casa com Titles da lista.")
        print("Titles disponíveis:", list(title_to_internal.keys()))
        return

    df_filtered = df[common_titles].copy()
    df_filtered.rename(columns=title_to_internal, inplace=True)

    # 9) INSERTs (linha a linha)
    ok = 0; erros = 0
    for i, row in df_filtered.iterrows():
        payload = {k: (None if (v == '' or v is None) else v) for k, v in row.to_dict().items()}
        success, resp = create_item(sess, SITE_URL, list_guid, payload)
        if success:
            ok += 1
            item_id = (resp.get('Id') or resp.get('ID') or "?") if isinstance(resp, dict) else "?"
            print(f"[{i}] Item criado. ID={item_id}")
        else:
            erros += 1
            print(f"[{i}] ERRO ao criar item: {resp}")

    print(f"\nResumo (INSERTs): Sucesso={ok} | Erros={erros}")

if __name__ == "__main__":
    main()
