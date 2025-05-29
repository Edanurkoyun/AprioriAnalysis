import pandas as pd
from collections import defaultdict
from itertools import combinations

# 1. Veri okuma ve işlem
df = pd.read_excel("C:\\Users\\edanu\\Desktop\\AprioriProject\\temizlenmis_veri2.xlsx", header=None)
urun_df = df.iloc[:, 1:]
transactions = urun_df.apply(lambda row: [str(item).strip().lower() for item in row.dropna()], axis=1).tolist()

# 2. Daha önce filtrelenen ürün listesi
filtered_products = pd.read_excel("C:\\Users\\edanu\\Desktop\\AprioriProject\\support_listesi_filtreli.xlsx")['Ürün']
filtered_products = filtered_products.str.lower().str.strip().tolist()

# 3. Sadece filtrelenmiş ürünleri içeren işlemler
filtered_transactions = []
for transaction in transactions:
    filtered_items = [item for item in transaction if item in filtered_products]
    if len(filtered_items) >= 2:
        filtered_transactions.append(filtered_items)

total_receipts = len(filtered_transactions)

# 4. İkili kombinasyonlar için sayıcılar
item_counts = defaultdict(int)
pair_counts = defaultdict(int)

for transaction in filtered_transactions:
    unique_items = set(transaction)
    for item in unique_items:
        item_counts[item] += 1
    for pair in combinations(unique_items, 2):
        sorted_pair = tuple(sorted(pair))
        pair_counts[sorted_pair] += 1

# 5. İkili kombinasyonlardan support, confidence hesapla ve support >= 0.05 filtrele
min_support = 0.05
filtered_pairs = {pair: count for pair, count in pair_counts.items() if count / total_receipts >= min_support}

# 6. İkili kombinasyonlarda geçen ürünlerin setini oluştur (sadece destekli ürünler)
products_in_filtered_pairs = set()
for pair in filtered_pairs.keys():
    products_in_filtered_pairs.update(pair)

# 7. Üçlü kombinasyonlar için yeni sayıcılar
triple_counts = defaultdict(int)

# 8. Üçlü kombinasyonları hesaplamak için filtrelenmiş ürünleri kullanarak transactions'ı yeniden filtrele
filtered_transactions_for_triples = []
for transaction in filtered_transactions:
    # Transaction'daki ürünlerden sadece products_in_filtered_pairs içindekiler alınacak
    filtered_items = [item for item in transaction if item in products_in_filtered_pairs]
    if len(filtered_items) >= 3:
        filtered_transactions_for_triples.append(filtered_items)

# 9. Üçlü kombinasyonları say
for transaction in filtered_transactions_for_triples:
    unique_items = set(transaction)
    for triple in combinations(unique_items, 3):
        sorted_triple = tuple(sorted(triple))
        triple_counts[sorted_triple] += 1

# 10. Üçlü kombinasyonların support ve confidence hesapla
data_triples = []
for triple, triple_support_count in triple_counts.items():
    support = triple_support_count / total_receipts
    item1, item2, item3 = triple

    # Confidence hesaplamak için her ürünün destek sayısını alalım
    conf1 = triple_support_count / item_counts[item1]
    conf2 = triple_support_count / item_counts[item2]
    conf3 = triple_support_count / item_counts[item3]

    data_triples.append({
        'Ürün 1': item1,
        'Ürün 2': item2,
        'Ürün 3': item3,
        'Support': round(support, 4),
        'Confidence (1→Diğerleri)': round(conf1, 4),
        'Confidence (2→Diğerleri)': round(conf2, 4),
        'Confidence (3→Diğerleri)': round(conf3, 4),
    })

# 11. DataFrame oluştur ve support'a göre sırala
result_triples_df = pd.DataFrame(data_triples)
result_triples_df = result_triples_df.sort_values(by='Support', ascending=False)

# 12. Sonucu Excel'e kaydet
result_triples_df.to_excel("C:\\Users\\edanu\\Desktop\\AprioriProject\\uclu_kurallar.xlsx", index=False)

print("Üçlü kombinasyonlar başarıyla hesaplandı ve Excel'e kaydedildi.")
