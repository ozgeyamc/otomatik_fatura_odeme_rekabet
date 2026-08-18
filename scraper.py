import requests
from bs4 import BeautifulSoup
import pandas as pd
from datetime import datetime
import os
import time

# İlettiğin banka listesi ve linkleri
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

# Bot olduğumuzu gizlemek için standart bir tarayıcı kimliği kullanıyoruz
HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8',
    'Accept-Language': 'tr-TR,tr;q=0.9,en-US;q=0.8,en;q=0.7',
}

def scrape_bank_data(bank_name, url):
    try:
        # Sitelerin bizi engellememesi için her istekte 2 saniye bekliyoruz
        time.sleep(2) 
        response = requests.get(url, headers=HEADERS, timeout=15)
        
        # Eğer site güvenlik duvarı bizi engellerse
        if response.status_code != 200:
            return "Güvenlik Duvarı Engeli", f"HTTP Hata Kodu: {response.status_code}"

        soup = BeautifulSoup(response.content, 'html.parser')
        
        # Sayfadaki tüm metni alıp küçük harfe çeviriyoruz
        text_content = soup.get_text(separator=' ', strip=True).lower()

        faiz_sonuc = "Faiz verisi tespit edilemedi"
        ucret_sonuc = "Ücret verisi tespit edilemedi"

        # Otomatik fatura anahtar kelimelerini arıyoruz
        if "otomatik fatura" in text_content or "fatura ödeme" in text_content:
            if "faiz" in text_content or "akdi" in text_content or "%" in text_content:
                faiz_sonuc = "Sayfada faiz verisi var (Detay için linke bakınız)"
            
            if "ücret" in text_content or "masraf" in text_content or "tl" in text_content:
                ucret_sonuc = "Sayfada ücret/masraf verisi var (Detay için linke bakınız)"

        return faiz_sonuc, ucret_sonuc

    except requests.exceptions.Timeout:
        return "Zaman Aşımı (Timeout)", "Site yanıt vermedi"
    except Exception as e:
        return "Bağlantı Hatası", "Manuel kontrol ediniz"

def main():
    print("Bankaların verileri çekiliyor...")
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

    # Veriyi Pandas DataFrame'e çevir
    df_new = pd.DataFrame(data_list)
    file_name = "OFO_Faiz_Oranlari.xlsx"
    
    # Eski Excel dosyası varsa verileri üstüne yazmadan alta ekle
    if os.path.exists(file_name):
        print("Mevcut Excel bulundu, yeni veriler altına ekleniyor...")
        df_existing = pd.read_excel(file_name)
        df_final = pd.concat([df_existing, df_new], ignore_index=True)
    else:
        print("Excel bulunamadı, sıfırdan oluşturuluyor...")
        df_final = df_new

    # Excel formatında kaydet
    df_final.to_excel(file_name, index=False)
    print(f"İşlem tamamlandı! Veriler {file_name} dosyasına yazıldı.")

if __name__ == "__main__":
    main()
