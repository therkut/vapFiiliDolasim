import os
import time
import requests
import pandas as pd
from datetime import datetime, timedelta
from zipfile import BadZipFile
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# === AYARLAR ===
DOWNLOAD_DIR = os.path.join(os.getcwd(), "data")
URL_FORM = "https://www.vap.org.tr/fiili-dolasim-raporu"
URL_API = "https://www.vap.org.tr/api/all-companies"
AS_FID = "5c19d6209383a7ba3c33b83b44390888f34db1f5"

CURL_HEADERS = {
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8",
    "Accept-Language": "en,tr;q=0.9",
    "Cache-Control": "max-age=0",
    "Connection": "keep-alive",
    "Sec-Fetch-Dest": "document",
    "Sec-Fetch-Mode": "navigate",
    "Sec-Fetch-Site": "none",
    "Sec-Fetch-User": "?1",
    "Sec-GPC": "1",
    "Upgrade-Insecure-Requests": "1",
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/141.0.0.0 Safari/537.36",
}

CURL_COOKIE_STRING = ("NSC_xxx.wbq.psh.us_mc=ffffffffc3a0da6845525d5f4f58455e445a4a42378b; "
                      "mkk_.vap.org.tr_%2F_wlf=AAAAAAXOjSz9clrR8aHBhHZZb-YxrYupvdvj1E0rHozq9YLltCswjrPnM3jL1oRY2xuQ0OczCKGgPklBsQDqc8mwW2Qq0O8tfu7WgmfJcVZhf8qKuw==&; "
                      "mkk_.vap.org.tr_%2F_wat=AAAAAAXhxC3oo1sRC8waFTx5yCFdfjxANZOGRtAoxzDVFNXRHPQ3sz5N-vqjCqSnlApGYnLjzdVZ5E21B1oVoR20Shv8#KLvoBmjrWpvNe4RhRlPST3r4kxkA&; "
                      "mkk=AAA78EruaDtDLAcAAAAAADt3BagVEO5sePNWO9pZS7XzUXovP54R7jK2tQw79gHoOw==vFHuaA==bVf3xObBJzUGVB2HECTttWJOmxU=")

def parse_cookie_string(cookie_str):
    return dict(part.strip().split("=", 1) for part in cookie_str.split(";") if "=" in part)

# === TARİH HESAPLAMA ===
def get_target_date():
    today = datetime.today()
    weekday = today.weekday()
    if weekday == 0:  # Pazartesi
        return today - timedelta(days=3)
    elif 1 <= weekday <= 4:  # Salı–Cuma
        return today - timedelta(days=1)
    else:
        return None

# === DOSYA İNDİRME (requests) ===
def try_post_download(date_str, download_dir):
    sess = requests.Session()
    sess.headers.update(CURL_HEADERS)
    sess.cookies.update(parse_cookie_string(CURL_COOKIE_STRING))

    data = {"date": date_str, "as_fid": AS_FID}
    print("🔁 requests POST:", URL_API, data)

    resp = sess.post(URL_API, data=data, timeout=30)
    print("HTTP", resp.status_code, "Content-Type:", resp.headers.get("Content-Type"))

    temp_path = os.path.join(download_dir, "temp_download.xlsx")
    with open(temp_path, "wb") as f:
        f.write(resp.content)

    # HTML hatası kontrolü
    if resp.content[:256].lstrip().startswith(b"<") or b"<html" in resp.content.lower():
        print("⚠️ Dönen içerik HTML — muhtemel hata sayfası.")
        return None

    try:
        pd.read_excel(temp_path, engine="openpyxl")
        print("✅ Excel başarılı şekilde indirildi:", temp_path)
        return temp_path
    except BadZipFile:
        print("⚠️ Geçerli Excel değil (BadZip).")
        return None

# === SELENIUM FALLBACK ===
def setup_selenium(download_dir):
    options = webdriver.ChromeOptions()
    options.add_experimental_option("prefs", {
        "download.default_directory": download_dir,
        "download.prompt_for_download": False,
        "directory_upgrade": True
    })
    options.add_argument("--headless=new")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    return webdriver.Chrome(options=options)

def wait_for_download(download_dir, extensions=(".xls", ".xlsx"), timeout=60):
    for _ in range(timeout):
        files = [f for f in os.listdir(download_dir) if f.endswith(extensions)]
        if files:
            return max(files, key=lambda f: os.path.getctime(os.path.join(download_dir, f)))
        time.sleep(1)
    return None

def selenium_fallback(date_str, download_dir):
    driver = setup_selenium(download_dir)
    try:
        driver.get(URL_FORM)
        WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.CSS_SELECTOR, "input#datepicker")))
        date_input = driver.find_element(By.CSS_SELECTOR, "input#datepicker")
        date_input.clear()
        date_input.send_keys(date_str)
        date_input.send_keys(Keys.ENTER)
        time.sleep(0.5)
        driver.find_element(By.CSS_SELECTOR, "input.submit-btn[value='Raporu Hazırla']").click()
        downloaded_file = wait_for_download(download_dir, timeout=60)
        if downloaded_file:
            return os.path.join(download_dir, downloaded_file)
        return None
    finally:
        driver.quit()

# === DOSYA İSİMLENDİRME (var olanın üzerine yazar) ===
def rename_files(excel_path):
    df = pd.read_excel(excel_path, nrows=5)
    date_candidates = [str(v) for v in df.values.flatten() if isinstance(v, str) and "/" in v]

    def safe_replace(src, dst):
        os.replace(src, dst)  # var olanın üzerine yazar

    if date_candidates:
        for d in date_candidates:
            try:
                date_str = datetime.strptime(d[:10], "%d/%m/%Y").strftime("%d-%m-%Y")
                new_excel = os.path.join(os.path.dirname(excel_path), f"Fiili_Dolasim_Raporu_MKK-{date_str}.xlsx")
                safe_replace(excel_path, new_excel)

                new_html = new_excel.replace(".xlsx", ".html")
                pd.read_excel(new_excel, engine="openpyxl").to_html(new_html, index=False, border=1, na_rep="")

                html_static = os.path.join(os.path.dirname(excel_path), "Fiili_Dolasim_Raporu_MKK.html")
                pd.read_excel(new_excel, engine="openpyxl").to_html(html_static, index=False, border=1, na_rep="")

                return new_excel, new_html, html_static
            except:
                continue

    # fallback
    fallback_excel = os.path.join(os.path.dirname(excel_path), "Fiili_Dolasim_Raporu_MKK.xlsx")
    safe_replace(excel_path, fallback_excel)

    fallback_html = fallback_excel.replace(".xlsx", ".html")
    pd.read_excel(fallback_excel, engine="openpyxl").to_html(fallback_html, index=False, border=1, na_rep="")

    html_static = os.path.join(os.path.dirname(excel_path), "Fiili_Dolasim_Raporu_MKK.html")
    pd.read_excel(fallback_excel, engine="openpyxl").to_html(html_static, index=False, border=1, na_rep="")

    return fallback_excel, fallback_html, html_static

# === ANA ===
def main():
    target_date = get_target_date()
    if target_date is None:
        print("🛑 Hafta sonu. Script çalışmayacak.")
        return

    os.makedirs(DOWNLOAD_DIR, exist_ok=True)
    date_str = target_date.strftime("%d/%m/%Y")
    print("📅 Hedef veri tarihi:", date_str)

    excel_path = try_post_download(date_str, DOWNLOAD_DIR)
    if not excel_path:
        print("⚠️ requests ile başarılı olunamadı; Selenium fallback çalıştırılıyor.")
        excel_path = selenium_fallback(date_str, DOWNLOAD_DIR)

    if not excel_path:
        print("❌ Her iki yöntem de başarısız oldu.")
        return

    excel_path, html_path, html_static = rename_files(excel_path)
    print(f"✅ Dosyalar oluşturuldu:\nExcel: {excel_path}\nHTML: {html_path}\nSabit HTML: {html_static}")

if __name__ == "__main__":
    main()
