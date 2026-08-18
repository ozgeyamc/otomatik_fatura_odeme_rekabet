import requests
from bs4 import BeautifulSoup
import pandas as pd
from datetime import datetime

def get_garanti_ofo_data():
    url = "https://www.garantibbva.com.tr/urun-ve-hizmet-ucretleri"
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
    }
    
    print("Garanti BBVA sayfası taranıyor...")
    
    try:
        response = requests.get(url, headers=headers, timeout=15)
        
        if response.status_code != 200:
            return "Engellendi (HTTP Hatası)", "-"

        soup = BeautifulSoup(response.content, 'html.parser')
        
        # 1. Sayfadaki tüm tabloları bul
        tables = soup.find_all('table')
        
        for table in tables:
            # 2. Sadece "otomatik fatura" geçen tabloya odaklan
            if 'otomatik fatura' in table.text.lower():
                rows = table.find_all('tr')
                
                for row in rows:
                    cells = row.find_all(['td', 'th'])
                    
                    # Eğer satırda hücre yoksa atla
                    if len(cells) < 2:
                        continue
                        
                    row_text = " ".join([cell.text.strip() for cell in cells]).lower()
                    
                    # 3. Kredi kartı ile ilgili satırı yakala
                    if 'kredi kartı' in row_text and ('faiz' in row_text or 'ücret' in row_text):
                        # Genellikle 1. hücre işlem adı, 2. veya 3. hücre orandır. 
                        # Sitenin yapısına göre buradaki indexleri [1] veya [2] olarak değiştirmek gerekebilir.
                        islem_adi = cells[0].text.strip()
                        deger = cells[1].text.strip() 
                        
                        return islem_adi, deger
                        
        return "Sayfada yapı değişmiş, veri bulunamadı", "-"
        
    except Exception as e:
        return f"Bağlantı Hatası: {e}", "-"

def main():
    today = datetime.today().strftime('%Y-%m-%d')
    islem_detayi, rakam = get_garanti_ofo_data()
    
    # Çekilen veriyi konsola yazdırarak kontrol edelim
    print("-" * 30)
    print(f"Bulunan İşlem: {islem_detayi}")
    print(f"Bulunan Rakam: {rakam}")
    print("-" * 30)
    
    # Excel'e yazma kısmı
    data = [{
        "Tarih": today,
        "Banka": "Garanti BBVA",
        "Kredi Kartı ile OFÖ Ödeme Faizi / Ücreti": rakam,
        "Anlık Ödeme Ücretleri": "Şimdilik test edilmedi",
        "Son Güncelleme": today
    }]
    
    df = pd.DataFrame(data)
    df.to_excel("Garanti_Test_OFO.xlsx", index=False)
    print("Test Excel'i oluşturuldu: Garanti_Test_OFO.xlsx")

if __name__ == "__main__":
    main()
