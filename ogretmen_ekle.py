import sqlite3

baglanti = sqlite3.connect("ktech.db")
imlec = baglanti.cursor()

# Test amaçlı bir öğretmen ekliyoruz
imlec.execute("""
    INSERT INTO Kullanicilar (eposta, sifre, ad_soyad, rol)
    VALUES ('ogretmen@kucukyali.edu.tr', '123456', 'Fahri Murat Terkeş', 'Öğretmen')
""")

baglanti.commit()
baglanti.close()
print("Test öğretmeni başarıyla eklendi!")