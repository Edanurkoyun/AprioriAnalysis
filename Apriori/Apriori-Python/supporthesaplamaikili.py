import pandas as pd  # Veri işleme için pandas kütüphanesi
from collections import defaultdict  # Sayım işlemleri için varsayılan sözlük yapısı
from itertools import combinations  # Kombinasyonlar oluşturmak için fonksiyon

# 1. Temizlenmiş Excel dosyasını okuyorum (header=None çünkü sütun başlıkları yok)
df = pd.read_excel("C:\\Users\\edanu\\Desktop\\AprioriProject\\temizlenmis_veri.xlsx", header=None)

# 2. İlk sütunu çıkarıyorum çünkü bu sütun fiş numaralarını içeriyor, analizde sadece ürün sütunları lazım
urun_df = df.iloc[:, 1:]

# 3. Her satırdaki ürünleri temizleyip liste haline getiriyorum:
#    - NaN değerleri çıkarıyorum,
#    - her ürünü string'e çeviriyorum,
#    - baş ve sondaki boşlukları siliyorum,
#    - tümünü küçük harfe çeviriyorum,
#    Böylece tüm fişleri temizlenmiş ürün listeleri olarak 'transactions' içine koyuyorum.
transactions = urun_df.apply(
    lambda row: [str(item).strip().lower() for item in row.dropna()], axis=1
).tolist()

# 4. Daha önce support değeri %3 ve üzeri olan ürünlerin listesini okuyorum
filtered_products = pd.read_excel("C:\\Users\\edanu\\Desktop\\AprioriProject\\support_listesi_filtreli.xlsx")['Ürün']

# 5. Ürün isimlerini küçük harfe çevirip baş/son boşluklarını temizliyorum, liste haline getiriyorum
filtered_products = filtered_products.str.lower().str.strip().tolist()

# 6. Support filtresinden geçmiş ürünleri içeren fişleri oluşturuyorum
filtered_transactions = []
for transaction in transactions:
    # Her fişte sadece filtrelenmiş ürünleri seçiyorum
    filtered_items = [item for item in transaction if item in filtered_products]
    # İkili ürün kombinasyonu için en az 2 ürün olmalı, öyleyse listeye ekliyorum
    if len(filtered_items) >= 2:
        filtered_transactions.append(filtered_items)

# 7. Filtrelenmiş fişlerin toplam sayısını hesaplıyorum
total_receipts = len(filtered_transactions)

# 8. Tekli ürünlerin kaç fişte göründüğünü saymak için sözlük oluşturuyorum (default 0)
item_counts = defaultdict(int)

# 9. İkili ürün kombinasyonlarının kaç fişte birlikte olduğunu saymak için sözlük oluşturuyorum
pair_counts = defaultdict(int)

# 10. Her fişi teker teker inceleyip destek sayacı hesaplıyorum
for transaction in filtered_transactions:
    # Aynı fişte bir ürün birden fazla ise tekrar saymamak için set'e çeviriyorum
    unique_items = set(transaction)

    # Tekli ürün destek sayısını artırıyorum
    for item in unique_items:
        item_counts[item] += 1

    # İkili ürün kombinasyonlarını oluşturuyorum (2'li tüm kombinasyonlar)
    for pair in combinations(unique_items, 2):
        # Çiftleri alfabetik sıraya göre sıralayıp tuple yapıyorum ki ('a','b') ve ('b','a') aynı sayılsın
        sorted_pair = tuple(sorted(pair))
        pair_counts[sorted_pair] += 1

# 11. Hesaplanan destek ve confidence değerlerini liste haline getiriyorum
data = []
for pair, pair_support_count in pair_counts.items():
    item1, item2 = pair

    # Support: bu ikilinin birlikte göründüğü fişlerin toplam fişlere oranı
    support = pair_support_count / total_receipts

    # Confidence (A→B): A ürününün bulunduğu fişlerin kaçında B de var
    confidence1 = pair_support_count / item_counts[item1]

    # Confidence (B→A): B ürününün bulunduğu fişlerin kaçında A da var
    confidence2 = pair_support_count / item_counts[item2]

    # Verileri sözlük olarak listeye ekliyorum
    data.append({
        'Ürün A': item1,
        'Ürün B': item2,
        'Support': round(support, 4),  # Dörde yuvarlıyorum
        'Confidence (A→B)': round(confidence1, 4),
        'Confidence (B→A)': round(confidence2, 4)
    })

# 12. Listeyi pandas DataFrame'e çeviriyorum
result_df = pd.DataFrame(data)

# 13. Support değerine göre büyükten küçüğe sıralıyorum
result_df = result_df.sort_values(by='Support', ascending=False)

# 14. Sonuçları Excel dosyasına kaydediyorum, indeks sütunu olmadan
result_df.to_excel("C:\\Users\\edanu\\Desktop\\AprioriProject\\ikili_kurallar.xlsx", index=False)

# 15. İşlem tamamlandığında ekrana bilgi mesajı yazdırıyorum
print("Filtrelenmiş verideki ikili kurallar başarıyla hesaplandı ve Excel'e kaydedildi.")
