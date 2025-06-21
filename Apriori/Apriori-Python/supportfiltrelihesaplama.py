import pandas as pd  # Pandas kütüphanesini veri işleme için içe aktar
from collections import defaultdict  # Ürün sayımı için varsayılan sözlük yapısı

# 1. Excel dosyasını oku (header=None: ilk satırı sütun ismi olarak alma)
df = pd.read_excel("C:\\Users\\edanu\\Desktop\\AprioriProject\\temizlenmis_veri.xlsx", header=None)

# 2. İlk sütunu çıkar (fiş numaraları olduğu için ürün sütunlarını alıyoruz)
urun_df = df.iloc[:, 1:]

# 3. Her satırdaki ürünleri temizle ve liste haline getir:
# - NaN olanları at,
# - Her değeri stringe çevir,
# - Başındaki ve sonundaki boşlukları temizle,
# - Küçük harfe dönüştür,
# Sonuçta her satır için ürün listesi elde et,
# Tüm satırları liste olarak 'transactions' içine koy.
transactions = urun_df.apply(
    lambda row: [str(item).strip().lower() for item in row.dropna()], axis=1
).tolist()

# 4. Ürünlerin kaç farklı fişte geçtiğini saymak için sözlük oluştur
# defaultdict(int) ile her yeni ürünün sayacı 0'dan başlar
item_counts = defaultdict(int)

# 5. Her fiş (transaction) için tekrarlayan ürünleri tek say, sonra ürün sayısını artır
for transaction in transactions:
    unique_items = set(transaction)  # Aynı fişte tekrar eden ürünleri bir kez saymak için set kullandık
    for item in unique_items:
        item_counts[item] += 1  # Ürünün geçtiği fiş sayısını artır

# 6. Toplam fiş sayısını hesapla (toplam transaction sayısı)
total_receipts = len(transactions)

# 7. Support değerlerini hesapla ve DataFrame oluştur
# Support = ürünün geçtiği fiş sayısı / toplam fiş sayısı
support_df = pd.DataFrame({
    'Ürün': list(item_counts.keys()),  # Ürün isimleri
    'Fiş Sayısı': list(item_counts.values()),  # Ürünün geçtiği fiş sayısı
    'Support': [count / total_receipts for count in item_counts.values()]  # Support oranları
})

# 8. Support değerine göre büyükten küçüğe sırala
support_df = support_df.sort_values(by='Support', ascending=False)

# 9. Minimum support değerini belirle (örnek: %3)
min_support = 0.03

# 10. Support değeri min_support'tan büyük veya eşit olan ürünleri filtrele
filtered_support_df = support_df[support_df['Support'] >= min_support]

# 11. Filtrelenmiş sonucu yeni Excel dosyasına kaydet
filtered_support_df.to_excel("C:\\Users\\edanu\\Desktop\\AprioriProject\\support_listesi_filtreli.xlsx", index=False)

# 12. İşlem tamamlandığında kullanıcıya bilgi ver
print("Support listesi başarıyla oluşturuldu ve Excel'e kaydedildi.")
