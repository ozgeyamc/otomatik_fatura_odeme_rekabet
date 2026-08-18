import requests
from bs4 import BeautifulSoup
import pandas as pd
from datetime import datetime
import os
import time
import re

BANKS = {
    "Akbank (Axess)": "https://www.axess.com.tr/axess/sayfa/1/369/faiz-ve-ucretler",
    "Garanti BBVA": "https://www.garantibbva.com.tr/urun-ve-hizmet-ucretleri",
    "Yapı Kredi": "https://www.yapikredi.com.tr/bireysel-bankacilik/hesaplama-araclari/bireysel-urun-ve-hizmet-ucretleri",
    "İş Bankası": "https://www.isbank.com.tr/urun-ve-hizmet-ucretleri",
    "DenizBank": "https://www.denizbonus.com/faiz-ve-ucretler",
    "QNB": "https://www.qnb.com.tr/yasal/urun-hizmet-ucretleri",
    "TEB": "https://www.teb.com.tr/urun-ve-hizmet-ucretleri/",
    "Halkbank": "https://www.halkbank.com.tr/tr/urun-ve-hizmet-ucretleri/kredi-kartlari-ve-banka-kartlari",
    "VakıfBank": "https://www.vakifkart.com.tr/ayricaliklar/firsatlar/otomatik-fatura-odeme",
    "Ziraat Bankası": "https://www.ziraatbank.com.tr/tr/urun-ve-hizmet-ucretleri"
}

HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8',
    'Accept-Language': 'tr-TR,tr;q=0.9,en-US;q=0.8,en;q=0.7',
}

def scrape_bank_data(bank_name, url):
    try:
        time.sleep(2) 
        response = requests.get(url, headers=HEADERS, timeout=15)
        
        if response.status_code != 200:
            return "Erişim Engeli", f"HTTP Kodu: {response.status_code}"

        soup = BeautifulSoup(response.content, 'html.parser')
        # Görünür metni alıyoruz
        text_content = soup.get_text(separator=' ', strip=True)

        # Regex ile Yüzdelik oranları bul (Örn: %3,50 veya 4.25%)
        faiz_oranlari = re.findall(r'%\s?\d+[.,]\d+|\d+[.,]\d+\s?%', text_content)
        
        # Regex ile TL tutarlarını bul (Örn: 40 TL, 15,50 TL)
        ucret_tutarlari = re.findall(r'\d+[.,]?\d*\s?[Tt][Ll]', text_content)

        # Aynı sayıların tekrarını önle ve sayfadaki ilk 3 farklı oranı al
        faizler_listesi = list(dict.fromkeys(faiz_oranlari))[:3]
        ucretler_listesi = list(dict.fromkeys(ucret_tutarlari))[:3]

        faiz_sonuc = " / ".join(faizler_listesi) if faizler_listesi else "Oran Bulunamadı"
        ucret_sonuc = " / ".join(ucretler_listesi) if ucretler_listesi else "Tutar Bulunamadı"

        return faiz_sonuc, ucret_sonuc

    except Exception:
        return "Bağlantı Hatası", "Bağlantı Hatası"

def main():
    print("Bankaların verileri çekiliyor, rakamlar ayıklanıyor...")
    today = datetime.today().strftime('%Y-%m-%d')
    data_list = []

    for bank, url in BANKS.items():
        print(f"{bank} taranıyor...")
        faiz_bilgisi, ucret_bilgisi = scrape_bank_data(bank, url)
        
        data_list.append({
            "Tarih": today,
            "Banka": bank,
            "Kredi Kartı ile OFÖ Ödeme Faizi": faiz_bilgisi,
            "Anlık Ödeme Ücretleri": ucret_bilgisi,
            "Son Güncelleme": today
        })

    df_new = pd.DataFrame(data_list)
    file_name = "OFO_Faiz_Oranlari.xlsx"
    
    if os.path.exists(file_name):
        df_existing = pd.read_excel(file_name)
        df_final = pd.concat([df_existing, df_new], ignore_index=True)
    else:
        df_final = df_new

    df_final.to_excel(file_name, index=False)
    print("İşlem tamamlandı! Excel oluşturuldu.")

if __name__ == "__main__":
    main()
