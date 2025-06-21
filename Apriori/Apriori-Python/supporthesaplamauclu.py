import pandas as pd  # Veri işlemleri için pandas kütüphanesi
from itertools import combinations  # Kombinasyon oluşturmak için fonksiyon
from collections import defaultdict  # Sayım işlemleri için varsayılan sözlük yapısı

# 1. Temizlenmiş Excel dosyasından transaction (fiş) verilerini okuyorum
#    (header=None: sütun başlıkları yok, ilk sütun fiş numarası, onu çıkarıyorum)
df = pd.read_excel("C:\\Users\\edanu\\Desktop\\AprioriProject\\temizlenmis_veri.xlsx", header=None)
urun_df = df.iloc[:, 1:]  # İlk sütun hariç tüm ürün sütunlarını seçiyorum

# 2. Her satırdaki ürünleri temizleyip liste haline getiriyorum:
#    - NaN değerleri çıkar,
#    - her ürünü stringe çevir,
#    - baş/son boşlukları sil,
#    - küçük harfe çevir,
# Böylece tüm fişleri ürün listeleri olarak 'transactions' içine koyuyorum
transactions = urun_df.apply(
    lambda row: [str(item).strip().lower() for item in row.dropna()], axis=1
).tolist()

# 3. Daha önce hesaplanmış ikili kombinasyonların support değerleriyle birlikte olduğu Excel dosyasını okuyorum
support_df = pd.read_excel("C:\\Users\\edanu\\Desktop\\AprioriProject\\ikili_kurallar.xlsx")

# 4. Support değeri %3 ve üzeri olan ikili kombinasyonları filtreliyorum
min_support = 0.03
filtered_pairs_df = support_df[support_df['Support'] >= min_support]

# 5. Bu filtrelenmiş ikili kombinasyonlardaki tüm ürünleri set olarak topluyorum
products_from_pairs = set()
filtered_pairs_set = set()  # Hızlı kontrol için ikilileri de set olarak saklıyorum
for _, row in filtered_pairs_df.iterrows():
    p1 = row['Ürün A'].strip().lower()
    p2 = row['Ürün B'].strip().lower()
    products_from_pairs.update([p1, p2])  # Ürünleri set'e ekle
    filtered_pairs_set.add(tuple(sorted([p1, p2])))  # İkiliyi sıralı tuple olarak ekle

# 6. Toplam transaction sayısını alıyorum (filtrelemeden önce)
total_receipts = len(transactions)

# 7. Transactionları sadece filtrelenmiş ürünleri içerecek şekilde filtreliyorum
filtered_transactions = []
for transaction in transactions:
    filtered_items = [item for item in transaction if item in products_from_pairs]
    # En az 3 ürün olmalı ki üçlü kombinasyon hesaplayabilelim
    if len(filtered_items) >= 3:
        filtered_transactions.append(filtered_items)

# 8. Üçlü kombinasyonların kaç kez göründüğünü saymak için sayıcı sözlüğü oluşturuyorum
triple_counts = defaultdict(int)

# 9. Filtrelenmiş transactionlarda üçlü kombinasyonları oluşturuyorum
#    Ancak bir üçlü kombinasyonun geçerli olması için,
#    üçlüyü oluşturan 3 ikili kombinasyonun hepsi filtrelenmiş ikililerde bulunmalı
for transaction in filtered_transactions:
    unique_items = set(transaction)  # Aynı fişteki ürünleri tekrar saymamak için set
    for triple in combinations(unique_items, 3):
        # Üçlünün altındaki 3 ikili kombinasyonu oluşturuyorum
        pairs_in_triple = [
            tuple(sorted([triple[0], triple[1]])),
            tuple(sorted([triple[0], triple[2]])),
            tuple(sorted([triple[1], triple[2]])),
        ]
        # Eğer bu 3 ikili kombinasyonun tamamı filtrelenmiş ikili kombinasyonlar içinde varsa,
        # üçlü kombinasyonun sayacını artırıyorum
        if all(pair in filtered_pairs_set for pair in pairs_in_triple):
            triple_counts[tuple(sorted(triple))] += 1

# 10. Her ürünün kaç fişte göründüğünü saymak için sayaç oluşturuyorum
item_counts = defaultdict(int)
for transaction in filtered_transactions:
    unique_items = set(transaction)
    for item in unique_items:
        item_counts[item] += 1

# 11. Üçlü kombinasyonların destek (support) ve güven (confidence) değerlerini hesaplıyorum
data_triples = []
for triple, count in triple_counts.items():
    support = count / total_receipts  # Destek: üçlünün toplam fişlere oranı
    item1, item2, item3 = triple

    # Confidence hesapları: üçlünün sayısı / ilgili ürünün tekil sayısı
    conf1 = count / item_counts[item1]
    conf2 = count / item_counts[item2]
    conf3 = count / item_counts[item3]

    # Sonuçları sözlük olarak listeye ekliyorum
    data_triples.append({
        'Ürün A': item1,
        'Ürün B': item2,
        'Ürün C': item3,
        'Support': round(support, 4),
        'Confidence (A→Diğerleri)': round(conf1, 4),
        'Confidence (B→Diğerleri)': round(conf2, 4),
        'Confidence (C→Diğerleri)': round(conf3, 4),
    })

# 12. Sonuç listesini DataFrame'e çevirip support değerine göre azalan şekilde sıralıyorum
result_triples_df = pd.DataFrame(data_triples)
result_triples_df = result_triples_df.sort_values(by='Support', ascending=False)

# 13. Sonucu Excel dosyasına kaydediyorum, indeks olmadan
result_triples_df.to_excel("C:\\Users\\edanu\\Desktop\\AprioriProject\\uclu_kurallar.xlsx", index=False)

# 14. İşlem tamamlandığında kullanıcıya bilgi veriyorum
print("Support değeri %3'ün üstündeki ikili kombinasyonlardan gelen ürünlerle üçlü kombinasyonlar hesaplandı ve kaydedildi.")
