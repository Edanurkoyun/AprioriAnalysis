import pandas as pd  # Veri işleme için pandas kütüphanesini içe aktarır
from collections import defaultdict  # Ürünleri sayaç gibi saklamak için defaultdict yapısını kullanır

# 1. Temizlenmiş veriyi Excel'den okur (her satır bir fiş, sütunlar ürünler)
df = pd.read_excel("C:\\Users\\edanu\\Desktop\\AprioriProject\\temizlenmis_veri.xlsx", header=None)

# 2. İlk sütunu (örneğin "Fiş-1", "Fiş-2" gibi) çıkar, sadece ürünlerin olduğu sütunları al
urun_df = df.iloc[:, 1:]

# 3. Her satırdaki ürünleri listeye dönüştür:
# - NaN (boş) değerleri atla
# - Her ürünü string yap
# - Başındaki ve sonundaki boşlukları temizle
# - Küçük harfe çevir
# Bu işlemi her satır için yap ve sonuçları `transactions` adlı listeye dönüştür
transactions = urun_df.apply(
    lambda row: [str(item).strip().lower() for item in row.dropna()], axis=1
).tolist()

# 4. Ürünleri saymak için bir defaultdict tanımlanır
# Her ürün, kaç farklı fişte geçiyorsa o kadar kez sayılacak (aynı fişte birden fazla kez varsa yine 1 kez sayılır)
item_counts = defaultdict(int)
for transaction in transactions:
    unique_items = set(transaction)  # Aynı fişte bir ürün birden fazla kez varsa sadece bir kez say
    for item in unique_items:
        item_counts[item] += 1  # Ürün o fişte geçtiği için sayacı 1 artır

# 5. Toplam fiş (satır) sayısını bul (yani kaç adet alışveriş yapıldı)
total_receipts = len(transactions)

# 6. Support değerlerini hesapla ve yeni bir DataFrame oluştur:
# - 'Ürün': Ürün ismi
# - 'Fiş Sayısı': Kaç fişte geçtiği
# - 'Support': Ürünün geçtiği fiş sayısının, toplam fiş sayısına oranı (destek değeri)
support_df = pd.DataFrame({
    'Ürün': list(item_counts.keys()),
    'Fiş Sayısı': list(item_counts.values()),
    'Support': [count / total_receipts for count in item_counts.values()]
})

# 7. Support değerine göre büyükten küçüğe sırala (en sık geçen ürünler en üstte olacak)
support_df = support_df.sort_values(by='Support', ascending=False)

# 8. Sonuçları yeni bir Excel dosyasına kaydet
support_df.to_excel("C:\\Users\\edanu\\Desktop\\AprioriProject\\support_listesi.xlsx", index=False)

# 9. İşlem tamamlandığında kullanıcıya bilgi ver
print("Support listesi başarıyla oluşturuldu ve Excel'e kaydedildi.")
