import pandas as pd  # Pandas kütüphanesini veri işleme için içe aktarır

# Temizlenmiş veri dosyasını Excel'den okur (üst bilgi satırı olmadığı için header=None)
df = pd.read_excel("C:\\Users\\edanu\\Desktop\\Apriori\\temizlenmis_veri.xlsx", header=None)

# İlk sütun (örneğin fiş numarası) gereksiz olduğu için onu çıkarıyoruz, sadece ürünlerin olduğu sütunları alıyoruz
urun_df = df.iloc[:, 1:] # iloc index numarası ile veri seçmeyi sağlıyor

# DataFrame'deki tüm ürün hücrelerini tek bir listeye indiriyoruz (satır/sütun ayrımını kaldırıyoruz)
all_items = urun_df.values.ravel() # ravel örneğin matris gibi tabloları tek bir listeye dönüştürür

# Bu listedeki:
# - NaN (boş) değerleri kaldır
# - Tüm elemanları string (yazı) yap
# - Başındaki ve sonundaki boşlukları temizle
# - Tüm harfleri küçük harfe çevir
cleaned_items = pd.Series(all_items).dropna().astype(str).str.strip().str.lower()

# Tekrar edenleri kaldırarak eşsiz ürün isimlerini elde ediyoruz
unique_items = cleaned_items.unique()

# Elde edilen eşsiz ürün listesini bir DataFrame'e çeviriyoruz (kolon adı: "Ürünler")
unique_df = pd.DataFrame(unique_items, columns=['Ürünler'])

# Eşsiz ürün listesini yeni bir Excel dosyasına kaydediyoruz (index yazma, sadece ürünler)
unique_df.to_excel("C:\\Users\\edanu\\Desktop\\Apriori\\esiz_urunler.xlsx", index=False)

# İşlem tamamlandıktan sonra kullanıcıya mesaj veriyoruz
print("Eşsiz ve temizlenmiş ürünler başarıyla yazıldı.")
