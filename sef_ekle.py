import sqlite3

baglanti = sqlite3.connect("ktech.db")
imlec = baglanti.cursor()

imlec.execute("""
    INSERT INTO Kullanicilar (eposta, sifre, ad_soyad, rol)
    VALUES ('sef@kucukyali.edu.tr', '123456', 'Ali Şef', 'Alan Şefi')
""")

baglanti.commit()
baglanti.close()
print("Test Alan Şefi başarıyla eklendi!")