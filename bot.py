import sys
import time
import random
import pyautogui
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import os

sys.stdout.reconfigure(encoding='utf-8')

# Excel dosyasını oku
excel_dosyasi = "Deneme.xlsx"
try:
    excel_veri = pd.read_excel(excel_dosyasi, header=0)
    excel_veri.columns = excel_veri.columns.str.strip()
    # Veriyi CSV'ye kaydetme
    csv_dosyasi = "google_sheets_data_filtered.csv"
    excel_veri.to_csv(csv_dosyasi, index=False)
    print(f"Veriler '{csv_dosyasi}' dosyasına kaydedildi.")

except FileNotFoundError:
    print(f"{excel_dosyasi} dosyası bulunamadı.")
    exit()

current_dir = os.path.dirname(os.path.abspath(__file__))
extension_path = os.path.join(current_dir, "pgojnojmmhpofjgdmaebadhbocahppod", "1.15.3_0")
user_data_dir = os.path.join(current_dir, "chrome_user_data")

chrome_options = Options()
chrome_options.add_argument(f"--load-extension={extension_path}")
chrome_options.add_argument(f"--user-data-dir={user_data_dir}")
chrome_options.add_argument("--profile-directory=Default")

service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service, options=chrome_options)

wait = WebDriverWait(driver, 10)

def kullanici_verisini_isle(kullanici_verisi):
    try:
        driver.get("https://support.google.com/legal/contact/lr_legalother?product=geo&hl=de&sjid=4014991289122781176-EU")
        time.sleep(4)

        # Cookies ekleme
        cookies = [
            {'name': 'test_cookie', 'value': 'value_test_cookie', 'domain': 'google.com'},
            {'name': 'user_pref', 'value': 'value_user_pref', 'domain': 'google.com'}
        ]
        for cookie in cookies:
            driver.add_cookie(cookie)

        # Almanya seçimi
        dropdown = wait.until(EC.element_to_be_clickable((By.CLASS_NAME, "sc-select"))).click()
        time.sleep(1)
        deutschland_seceneği = wait.until(EC.element_to_be_clickable((By.XPATH, "//li[text()='Deutschland']"))).click()
        time.sleep(1)

        # Kullanıcı verisi form alanlarını doldurma
        alanlar = [
            ("full_name", kullanici_verisi.get("Vollstandiger_name")),
            ("companyname", kullanici_verisi.get("Name des Unternehmens")),
            ("representedrightsholder", kullanici_verisi.get("Name des Unternehmens")),
            ("contact_email_noprefill", kullanici_verisi.get("Mail")),
            ("url_box3_geo_germany", kullanici_verisi.get("Link")),
            ("legalother_explain_googlemybusiness_not_germany", kullanici_verisi.get("Text")),
            ("legalother_quote", kullanici_verisi.get("Unnamed: 0")),
            ("legalother_quote_googlemybusiness_not_germany", kullanici_verisi.get("Unnamed: 0")),
            ("signature", kullanici_verisi.get("Vollstandiger_name")),
        ]

        for alan_id, deger in alanlar:
            if deger:
                alan = wait.until(EC.element_to_be_clickable((By.ID, alan_id)))
                alan.clear()
                alan.send_keys(deger)
                time.sleep(1)

        # İkinci URL alanı varsa ekleyin
        link_2 = kullanici_verisi.get("Link 2")
        if link_2 and not pd.isna(link_2):  # "Link 2" değeri None, NaN veya boş değilse
            onay_kutusu = wait.until(EC.element_to_be_clickable((By.ID, "add_another_url_checkbox--add")))
            driver.execute_script("arguments[0].click();", onay_kutusu)
            yeni_url_alani = wait.until(EC.element_to_be_clickable((By.ID, "url_box3_googlemybusiness_2")))
            yeni_url_alani.clear()
            yeni_url_alani.send_keys(link_2)
            time.sleep(1)
        else:
            print("İkinci URL yok veya geçersiz, 'Weiteres Feld hinzufügen' tuşuna basılmadı.")


        # Radyo butonları ve onay kutuları
        nein_radyo_butonu = wait.until(EC.element_to_be_clickable((By.ID, "is_geo_ugc_imagery--no")))
        driver.execute_script("arguments[0].click();", nein_radyo_butonu)
        time.sleep(1)

        onay_checkbox = wait.until(EC.element_to_be_clickable((By.ID, "legal_consent_statement--agree")))
        driver.execute_script("arguments[0].click();", onay_checkbox)
        time.sleep(5)

        try:
            gonder_butonu = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[text()='Senden']")))
            driver.execute_script("arguments[0].click();", gonder_butonu)
            print("Gönder butonuna başarıyla tıklandı.")
            time.sleep(70)

            driver.execute_script("arguments[0].click();", gonder_butonu)
            print("Gönder butonuna ikinci kez başarıyla tıklandı.")
            time.sleep(10)

        except Exception as e:
            print(f"'Senden' butonuna tıklanamadı: {e}")

    except Exception as e:
        print(f"Bir hata oluştu: {e}")
        return False

csv_dosyasi = "google_sheets_data_filtered.csv"
veri = pd.read_csv(csv_dosyasi)
veri["H"] = veri["H"].astype(str)
excel_veri["H"] = excel_veri["H"].astype(str) 

for index, satir in veri.iterrows():
    if satir["H"] != "tamamlandı": 
        kullanici_verisi = satir.to_dict()
        basarili = kullanici_verisini_isle(kullanici_verisi)
        basarili= True
        if basarili:
            veri.at[index, "H"] = "tamamlandı" 
            excel_veri.at[index, "H"] = "tamamlandı" 
            try:
                veri.to_csv(csv_dosyasi, index=False)  
                excel_veri.to_excel(excel_dosyasi, index=False)
                print(f"Veriler başarıyla kaydedildi. Kullanıcı: {satir['Vollstandiger_name']}")
            except Exception as e:
                print(f"Dosyalar kaydedilirken bir hata oluştu: {e}")
        else:
            print(f"İşlem başarısız: {satir['Vollstandiger_name']}")

veri.to_csv(csv_dosyasi, index=False)  
excel_veri.to_excel(excel_dosyasi, index=False)
print("Tüm kullanıcılar işlendi. Uygulama kapatılıyor...")
driver.quit()