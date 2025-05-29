import pandas as pd

# Excel dosyasını oku
df = pd.read_excel("C:\\Users\\edanu\\Desktop\\AprioriProject\\temizlenmis_veri2.xlsx", header=None)

# İlk sütunu çıkar
urun_df = df.iloc[:, 1:]

# Tüm ürünleri tek listeye indir
all_items = urun_df.values.ravel()

# NaN'leri kaldır, string'e çevir, küçük harf yap, strip (baş/son boşlukları temizle)
cleaned_items = pd.Series(all_items).dropna().astype(str).str.strip().str.lower()

# Eşsiz ürünleri al
unique_items = cleaned_items.unique()

# DataFrame'e çevir
unique_df = pd.DataFrame(unique_items, columns=['Ürünler'])

# Excel dosyasına yaz
unique_df.to_excel("C:\\Users\\edanu\\Desktop\\AprioriProject\\esiz_urunler.xlsx", index=False)

print("Eşsiz ve temizlenmiş ürünler başarıyla yazıldı.")

