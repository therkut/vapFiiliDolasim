import os, time, pandas as pd
from selenium import webdriver
from datetime import datetime, timedelta
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# === AYARLAR ===
DOWNLOAD_DIR = os.path.join(os.getcwd(), "data")
os.makedirs(DOWNLOAD_DIR, exist_ok=True)
URL_FORM = "https://www.vap.org.tr/api/all-companies"

TEST = False              # 🧪 Test modu aktif/pasif
TEST_OFFSET_DAYS = 3     # 🔧 Test modunda kaç gün geriye gidecek

# === TARİH HESAPLAMA ===
def get_target_date():
    today = datetime.today()
    weekday = today.weekday()

    # 🧪 Test modundaysa hafta sonu kısıtı yok
    if TEST:
        target = today - timedelta(days=TEST_OFFSET_DAYS)
        return target

    # Normal mod (hafta sonu hariç)
    if weekday == 0:  # Pazartesi -> önceki Cuma
        target = today - timedelta(days=3)
    elif 1 <= weekday <= 4:  # Salı-Cuma -> önceki gün
        target = today - timedelta(days=1)
    else:
        return None  # Hafta sonu

    return target

# === KLASÖR TEMİZLEME ===
def clear_old_downloads(download_dir):
    for f in os.listdir(download_dir):
        if f.endswith((".xls", ".xlsx", ".crdownload")):
            try:
                os.remove(os.path.join(download_dir, f))
            except Exception as e:
                print(f"⚠️ {f} silinemedi: {e}")

# === SELENIUM SETUP ===
def setup_driver(download_dir):
    options = webdriver.ChromeOptions()
    options.add_experimental_option("prefs", {
        "download.default_directory": download_dir,
        "download.prompt_for_download": False,
        "directory_upgrade": True,
        "safebrowsing.enabled": True
    })
    options.add_argument("--headless")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--window-size=1920,1080")
    return webdriver.Chrome(options=options)

# === DOSYA İNDİRME BEKLEME ===
def wait_for_download(download_dir, timeout=20):
    end_time = time.time() + timeout
    while time.time() < end_time:
        files = [f for f in os.listdir(download_dir) if f.endswith((".xls", ".xlsx"))]
        if files:
            latest = max(files, key=lambda f: os.path.getmtime(os.path.join(download_dir, f)))
            full_path = os.path.join(download_dir, latest)
            if not latest.endswith(".crdownload"):
                size1 = os.path.getsize(full_path)
                time.sleep(1)
                size2 = os.path.getsize(full_path)
                if size1 == size2:
                    return full_path
        time.sleep(1)
    return None

# === EXCEL İNDİRME ===
def download_excel(date_str):
    clear_old_downloads(DOWNLOAD_DIR)
    driver = setup_driver(DOWNLOAD_DIR)
    try:
        driver.get(URL_FORM)
        WebDriverWait(driver, 40).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "form.all-companies-form"))
        )
        date_input = WebDriverWait(driver, 30).until(
            EC.visibility_of_element_located((By.ID, "datepicker"))
        )
        date_input.clear()
        date_input.send_keys(date_str)
        submit_btn = driver.find_element(By.CSS_SELECTOR, "input.submit-btn")
        driver.execute_script("arguments[0].click();", submit_btn)

        file_path = wait_for_download(DOWNLOAD_DIR, timeout=20)
        if file_path is None:
            raise Exception("Excel dosyası indirilemedi veya timeout oluştu.")
        return file_path
    finally:
        driver.quit()

# === EXCEL → HTML ===
def excel_to_html(excel_path, target_date_str):
    date_suffix = target_date_str.replace("/", "-")
    new_excel = os.path.join(DOWNLOAD_DIR, f"Fiili_Dolasim_Raporu_MKK-{date_suffix}.xlsx")
    os.replace(excel_path, new_excel)

    html_path = new_excel.replace(".xlsx", ".html")
    html_static = os.path.join(DOWNLOAD_DIR, "Fiili_Dolasim_Raporu_MKK.html")

    df = pd.read_excel(new_excel, engine="openpyxl")
    df.to_html(html_path, index=False, border=1, na_rep="")
    df.to_html(html_static, index=False, border=1, na_rep="")

    return new_excel, html_path, html_static

# === ANA ===
if __name__ == "__main__":
    target_date = get_target_date()

    if target_date is None:
        print("🛑 Hafta sonu. Script çalışmayacak.")
    else:
        date_str = target_date.strftime("%d/%m/%Y")
        mode_info = f"(TEST, {TEST_OFFSET_DAYS} gün geri)" if TEST else ""
        print(f"📅 Hedef veri tarihi: {date_str} {mode_info}")
        try:
            excel_file = download_excel(date_str)
            excel_path, html_path, html_static = excel_to_html(excel_file, date_str)
            print(f"✅ Dosyalar hazır:\nExcel: {excel_path}\nHTML: {html_path}\nSabit HTML: {html_static}")
        except Exception as e:
            print(f"❌ Hata oluştu: {e}")
