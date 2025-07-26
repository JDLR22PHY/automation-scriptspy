"""
Automated Document Downloader from Government Platform
Author: Juan David Lamus Rincón

This script logs into a secure platform, searches for specific codes,
downloads related documents, and organizes them into folders.

Note: Sensitive data like username, password, and platform URL
must be configured manually before use.
"""

import os
import re
import time
import shutil
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains

# --- USER CONFIGURATION ---
USERNAME = "your_username_here"
PASSWORD = "your_password_here"
CHROMEDRIVER_PATH = "/usr/bin/chromedriver"
PLATFORM_URL = "http://your.platform.url/login"

EXCEL_FILE = "codes.xlsx"  # Input Excel file
TEMP_DOWNLOAD_DIR = os.path.join(os.getcwd(), "TempDownloads")
FINAL_DOWNLOAD_DIR = os.path.join(os.getcwd(), "Downloads")

# --- Read codes from Excel ---
df = pd.read_excel(EXCEL_FILE, header=None, dtype=str)
raw_codes = df.iloc[:, 0].astype(str).tolist()
code_list = []

for cell in raw_codes:
    cell = cell.strip()
    if not cell or cell[0].isalpha():
        continue
    parts = re.split(r'[,\s]+', cell)
    for code in parts:
        match = re.match(r'^\d{4}[A-Z]{2}\d{5}', code)
        if match:
            code_list.append(match.group(0))

# --- Prepare Chrome and WebDriver ---
if not os.path.exists(TEMP_DOWNLOAD_DIR):
    os.makedirs(TEMP_DOWNLOAD_DIR)

chrome_options = Options()
chrome_options.add_experimental_option("prefs", {
    "download.default_directory": TEMP_DOWNLOAD_DIR,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True
})
driver = webdriver.Chrome(service=Service(CHROMEDRIVER_PATH), options=chrome_options)
driver.get(PLATFORM_URL)
time.sleep(3)

# --- Log in ---
driver.find_element(By.ID, "lgIniciarSesion_UserName").send_keys(USERNAME)
driver.find_element(By.ID, "lgIniciarSesion_Password").send_keys(PASSWORD + Keys.RETURN)
time.sleep(5)

# --- Navigate platform ---
WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, "//span[text()='Archivo de Correspondencia']"))
).click()

WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, "//a[contains(@id,'lnkArchivoTodo')]"))
).click()

main_window = driver.current_window_handle
WebDriverWait(driver, 10).until(lambda d: len(d.window_handles) > 1)

for handle in driver.window_handles:
    if handle != main_window:
        driver.switch_to.window(handle)
        break

# --- Process each code ---
for code in code_list:
    driver.refresh()

    # Clear dates
    WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.ID, "fraCriteriosCaracteristicas_deElaboradoDesde_dateInput"))
    ).clear()
    WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.ID, "fraCriteriosCaracteristicas_deElaboradoHasta_dateInput"))
    ).clear()

    input_field = driver.find_element(By.ID, "fraCriteriosCaracteristicas_txtCodigo")

    # Prepare folder
    destination_folder = os.path.join(FINAL_DOWNLOAD_DIR, code)
    if not os.path.exists(destination_folder):
        os.makedirs(destination_folder)

    input_field.send_keys(code)
    driver.find_element(By.ID, "btnAplicarFiltro_input").click()

    # Wait for table and double-click
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "gRegistros_ctl00"))
    )
    actions = ActionChains(driver)
    table = driver.find_element(By.ID, "gRegistros_ctl00")
    actions.double_click(table).perform()

    WebDriverWait(driver, 10).until(lambda d: len(d.window_handles) > 1)
    for handle in driver.window_handles:
        if handle != main_window:
            driver.switch_to.window(handle)
            break

    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.TAG_NAME, "iframe"))
    )
    driver.switch_to.frame(driver.find_element(By.TAG_NAME, "iframe"))

    WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//span[text()='Documentos']"))
    ).click()

    doc_table = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "vdirDOCUMENTOSfraDocumentosascx_gRegistros_ctl00"))
    )
    rows = doc_table.find_elements(By.TAG_NAME, "tr")[1:]

    for row in rows:
        try:
            actions.double_click(row).perform()
            time.sleep(4)
        except:
            pass

    WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.ID, "barraAceptarCancelar_btnAceptar_input"))
    ).click()
    driver.switch_to.default_content()

    for handle in driver.window_handles:
        if handle != main_window:
            driver.switch_to.window(handle)
            break

    input_field = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.ID, "fraCriteriosCaracteristicas_txtCodigo"))
    )
    input_field.clear()

    # Move downloaded files
    time.sleep(5)
    for filename in os.listdir(TEMP_DOWNLOAD_DIR):
        shutil.move(os.path.join(TEMP_DOWNLOAD_DIR, filename), os.path.join(destination_folder, filename))

input("Press Enter to exit...")

