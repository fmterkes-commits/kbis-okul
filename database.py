import sqlite3
import os

def veritabani_kur():
    if os.path.exists("ktech.db"):
        try:
            os.remove("ktech.db")
            print("Eski veritabanı başarıyla silindi.")
        except Exception as e:
            print("HATA: Veritabanı silinemedi! Lütfen Uvicorn sunucusunu durdurduğunuzdan emin olun.")
            return

    baglanti = sqlite3.connect("ktech.db")
    imlec = baglanti.cursor()

    imlec.execute("""
    CREATE TABLE IF NOT EXISTS Kullanicilar (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        eposta TEXT UNIQUE NOT NULL,
        sifre TEXT NOT NULL,
        ad_soyad TEXT NOT NULL,
        rol TEXT NOT NULL 
    )
    """)

    imlec.execute("""
    CREATE TABLE IF NOT EXISTS Mekanlar (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        bina_adi TEXT NOT NULL,
        kat TEXT NOT NULL,
        alan_adi TEXT NOT NULL
    )
    """)

    # YENİLİK: atama_tarihi ve sef_kapanis_tarihi sütunları eklendi!
    imlec.execute("""
    CREATE TABLE IF NOT EXISTS Bildirimler (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        bildiren_id INTEGER NOT NULL,
        mekan_id INTEGER NOT NULL,
        tur TEXT NOT NULL,           
        aciliyet INTEGER NOT NULL,   
        aciklama TEXT NOT NULL,
        fotograf_yolu TEXT,          
        durum TEXT DEFAULT 'Yeni',   
        
        atanan_sef_id INTEGER,       
        sef_ekibi TEXT,              
        sef_cozum_notu TEXT,         
        cozum_fotografi TEXT,        
        idareci_kapanis_notu TEXT,   
        
        tarih TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        atama_tarihi TIMESTAMP,
        sef_kapanis_tarihi TIMESTAMP,
        kapanis_tarihi TIMESTAMP     
    )
    """)

    imlec.execute("""
        INSERT INTO Kullanicilar (eposta, sifre, ad_soyad, rol)
        VALUES ('admin@kucukyali.edu.tr', '123456', 'Teknik Müdür Yardımcısı', 'İdareci')
    """)

    baglanti.commit()
    baglanti.close()
    print("Mükemmel! K-BİS veritabanı tam istediğiniz iş akışına göre güncellendi.")

if __name__ == "__main__":
    veritabani_kur()