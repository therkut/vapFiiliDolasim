import os
import time
import pandas as pd
import requests
from datetime import datetime, timedelta
from pathlib import Path

# Selenium yalnızca fallback için
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# === AYARLAR ===
DOWNLOAD_DIR = os.path.join(os.getcwd(), "data")
URL = "https://www.vap.org.tr/api/all-companies"
AS_FID = "5c19d6209383a7ba3c33b83b44390888f34db1f5"

# === TARİH BELİRLEME ===
def get_target_date():
    today = datetime.today()
    weekday = today.weekday()
    if weekday == 0:  # Pazartesi
        return today - timedelta(days=3)
    elif 1 <= weekday <= 4:  # Salı–Cuma
        return today - timedelta(days=1)
    else:
        return None

# === EXCEL → HTML DÖNÜŞÜM ===
def convert_excel_to_html(excel_path, html_path):
    try:
        df = pd.read_excel(excel_path, engine="openpyxl")  # .xlsx
    except Exception:
        df = pd.read_excel(excel_path, engine="xlrd")      # .xls fallback
    df.to_html(html_path, index=False, border=1, na_rep="")

# === SELENIUM SETUP ===
def setup_selenium(download_dir):
    options = webdriver.ChromeOptions()
    prefs = {
        "download.default_directory": download_dir,
        "download.prompt_for_download": False,
        "directory_upgrade": True
    }
    options.add_experimental_option("prefs", prefs)
    options.add_argument("--headless=new")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    return webdriver.Chrome(options=options)

# === DOSYA İNDİRME BEKLEME ===
def wait_for_download(download_dir, extensions=(".xls", ".xlsx"), timeout=30):
    for _ in range(timeout):
        files = [f for f in os.listdir(download_dir) if f.endswith(extensions)]
        if files:
            return max(files, key=lambda f: os.path.getctime(os.path.join(download_dir, f)))
        time.sleep(1)
    return None

# === ANA FONKSİYON ===
def main():
    target_date = get_target_date()
    if target_date is None:
        print("🛑 Hafta sonu. Script çalışmayacak.")
        return

    os.makedirs(DOWNLOAD_DIR, exist_ok=True)
    date_str = target_date.strftime("%d/%m/%Y")
    xls_name = f"Fiili_Dolasim_Raporu_MKK-{target_date.strftime('%d-%m-%Y')}.xlsx"
    xls_path = os.path.join(DOWNLOAD_DIR, xls_name)

    print(f"📅 Hedef veri tarihi: {date_str}")

    # === 1️⃣ POST ile indirme denemesi ===
    try:
        print("🌐 Excel POST request ile indiriliyor...")
        r = requests.post(URL, data={"date": date_str, "as_fid": AS_FID})
        r.raise_for_status()
        temp_path = os.path.join(DOWNLOAD_DIR, "temp_download.xlsx")
        with open(temp_path, "wb") as f:
            f.write(r.content)
        try:
            pd.read_excel(temp_path, engine="openpyxl")
            os.rename(temp_path, xls_path)
            print(f"📄 İndirilen dosya (POST): {xls_name}")
        except Exception:
            print("⚠️ POST ile indirilen dosya geçersiz. Selenium fallback başlatılıyor...")
            os.remove(temp_path)
            raise Exception("Invalid Excel format")
    except Exception:
        # === 2️⃣ Selenium fallback ===
        driver = setup_selenium(DOWNLOAD_DIR)
        try:
            driver.get(URL)
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "input#datepicker"))
            )
            driver.execute_script(f"""
                var input = document.querySelector("#datepicker");
                input.value = '{date_str}';
                input.dispatchEvent(new Event('change', {{ bubbles: true }}));
            """)
            driver.find_element(By.CSS_SELECTOR, "input.submit-btn").click()
            downloaded_file = wait_for_download(DOWNLOAD_DIR)
            if not downloaded_file:
                print("❌ Selenium ile dosya indirilemedi.")
                return
            os.rename(os.path.join(DOWNLOAD_DIR, downloaded_file), xls_path)
            print(f"📄 İndirilen dosya (Selenium): {xls_name}")
        finally:
            driver.quit()

    # === HTML DÖNÜŞÜM ===
    dated_html_path = xls_path.replace(".xlsx", ".html")
    fixed_html_path = os.path.join(DOWNLOAD_DIR, "Fiili_Dolasim_Raporu_MKK.html")

    print("🔄 HTML'e dönüştürülüyor...")
    try:
        convert_excel_to_html(xls_path, dated_html_path)
        convert_excel_to_html(xls_path, fixed_html_path)
        print(f"✅ HTML dosyaları oluşturuldu:")
        print(f"  - {os.path.basename(dated_html_path)}")
        print(f"  - {os.path.basename(fixed_html_path)}")
    except Exception as e:
        print(f"❌ HTML dönüşümünde hata: {e}")

if __name__ == "__main__":
    main()
