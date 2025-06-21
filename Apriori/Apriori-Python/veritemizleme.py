import pandas as pd

# Excel dosyasının bilgisayardaki tam yolunu belirtiyoruz
dosya_yolu = "C:\\Users\\edanu\\Desktop\\AprioriProject\\AprioriProject.xlsx"  # Excel dosyasının konumu

# Excel dosyasını pandas ile okuyoruz ve bir DataFrame'e yüklüyoruz
df = pd.read_excel(dosya_yolu)

# İlk sütunu (örneğin "Fiş No") DataFrame'in satır index’i olarak ayarlıyoruz
df = df.set_index(df.columns[0])

# Her satır için:
# - NaN (boş) hücreleri kaldır,
# - Geriye kalan değerleri bir listeye dönüştür,
# - Tüm satırları bu şekilde işle ve sonucu df_cleaned'e ata
df_cleaned = df.apply(lambda x: x.dropna().tolist(), axis=1)

# Yukarıdaki listeleri hücrelere açarak yeni bir DataFrame oluşturuyoruz:
# - Her listedeki eleman (ürün) ayrı bir sütuna yerleştirilir
# - Eski fiş numaraları (index) korunur
df_expanded = pd.DataFrame(df_cleaned.tolist(), index=df_cleaned.index)

# Temizlenmiş ve açılmış verileri yeni bir Excel dosyasına yazdırıyoruz
# - index=True: Fiş numaraları yazılsın
# - header=False: Üst satıra sütun isimleri yazılmasın
df_expanded.to_excel("C:\\Users\\edanu\\Desktop\\Apriori\\temizlenmis_veri.xlsx", index=True, header=False)