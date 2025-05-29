import pandas as pd
from collections import defaultdict
from itertools import combinations

# 1. Temizlenmiş veriyi oku
df = pd.read_excel("C:\\Users\\edanu\\Desktop\\AprioriProject\\temizlenmis_veri2.xlsx", header=None)

# 2. İlk sütunu çıkar
urun_df = df.iloc[:, 1:]

# 3. Tüm fişleri hazırla (küçük harfe çevir + boşluk temizle)
transactions = urun_df.apply(lambda row: [str(item).strip().lower() for item in row.dropna()], axis=1).tolist()

# 4. Daha önce filtrelediğin (support >= %3) ürün listesini al
filtered_products = pd.read_excel("C:\\Users\\edanu\\Desktop\\AprioriProject\\support_listesi_filtreli.xlsx")['Ürün']
filtered_products = filtered_products.str.lower().str.strip().tolist()

# 5. Sadece bu ürünleri içeren fişleri oluştur
filtered_transactions = []
for transaction in transactions:
    filtered_items = [item for item in transaction if item in filtered_products]
    if len(filtered_items) >= 2:  # ikili kombinasyon için en az 2 ürün gerek
        filtered_transactions.append(filtered_items)

# 6. Toplam fiş sayısı (filtrelenmiş)
total_receipts = len(filtered_transactions)

# 7. Tekli ürün destek sayacı
item_counts = defaultdict(int)
# 8. İkili ürün destek sayacı
pair_counts = defaultdict(int)

# 9. Her fiş için ürün desteklerini ve kombinasyonları hesapla
for transaction in filtered_transactions:
    unique_items = set(transaction)
    
    # Tekli ürün destek
    for item in unique_items:
        item_counts[item] += 1

    # İkili kombinasyonlar
    for pair in combinations(unique_items, 2):
        sorted_pair = tuple(sorted(pair))  # sıralı olsun
        pair_counts[sorted_pair] += 1

# 10. Sonuçları hesapla ve listele
data = []
for pair, pair_support_count in pair_counts.items():
    item1, item2 = pair
    support = pair_support_count / total_receipts
    
    confidence1 = pair_support_count / item_counts[item1]
    confidence2 = pair_support_count / item_counts[item2]
    
    data.append({
        'Ürün A': item1,
        'Ürün B': item2,
        'Support': round(support, 4),
        'Confidence (A→B)': round(confidence1, 4),
        'Confidence (B→A)': round(confidence2, 4)
    })

# 11. Sonuçları DataFrame'e aktar
result_df = pd.DataFrame(data)
result_df = result_df.sort_values(by='Support', ascending=False)


# 12. Excel'e kaydet
result_df.to_excel("C:\\Users\\edanu\\Desktop\\AprioriProject\\ikili_kurallar.xlsx", index=False)

print("Filtrelenmiş verideki ikili kurallar başarıyla hesaplandı ve Excel'e kaydedildi.")

