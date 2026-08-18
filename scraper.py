import requests
from bs4 import BeautifulSoup
import pandas as pd
from datetime import datetime
import os

# Sitelerin bot korumalarına takılmamak için tarayıcı taklidi yapıyoruz
HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36'
}

def get_bank_data():
    # Burası ana veri toplama yerin. Copilot ile her banka için ayrı fonksiyonlar yazıp burada birleştirebilirsin.
    today = datetime.today().strftime('%Y-%m-%d')
    data_list = []
    
    # Örnek 1: Garanti BBVA (Burayı Copilot ile request atıp BeautifulSoup ile çekecek şekilde genişletirsin)
    data_list.append({
        "Tarih": today,
        "Banka": "Garanti",
        "Kredi Kartı ile OFÖ Ödeme Faizi": "Web'den çekilen veri", 
        "Anlık Ödeme Ücretleri": "Web'den çekilen veri",
        "Son Güncelleme": today
    })
    
    # Örnek 2: Akbank
    data_list.append({
        "Tarih": today,
        "Banka": "Akbank",
        "Kredi Kartı ile OFÖ Ödeme Faizi": "Web'den çekilen veri", 
        "Anlık Ödeme Ücretleri": "Web'den çekilen veri",
        "Son Güncelleme": today
    })
    
    return data_list

def main():
    print("Veriler çekiliyor...")
    new_data = get_bank_data()
    df_new = pd.DataFrame(new_data)
    
    file_name = "OFO_Faiz_Oranlari.xlsx"
    
    # Eğer Excel zaten varsa, eski verileri koruyup yenileri altına ekliyoruz
    if os.path.exists(file_name):
        print("Mevcut Excel bulundu, yeni veriler ekleniyor...")
        df_existing = pd.read_excel(file_name)
        df_final = pd.concat([df_existing, df_new], ignore_index=True)
    else:
        print("Excel bulunamadı, sıfırdan oluşturuluyor...")
        df_final = df_new

    # Excel'e kaydet
    df_final.to_excel(file_name, index=False)
    print(f"İşlem tamamlandı. Veriler {file_name} dosyasına yazıldı.")

if __name__ == "__main__":
    main()
