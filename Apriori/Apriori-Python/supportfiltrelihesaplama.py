import pandas as pd
from collections import defaultdict

# 1. Excel dosyasını okuyor
df = pd.read_excel("C:\\Users\\edanu\\Desktop\\AprioriProject\\temizlenmis_veri2.xlsx", header=None)


urun_df = df.iloc[:, 1:]


transactions = urun_df.apply(lambda row: [str(item).strip().lower() for item in row.dropna()], axis=1).tolist()

# 4. Ürünleri say (aynı fişte tekrar eden ürünleri bir kez say)
item_counts = defaultdict(int)
for transaction in transactions:
    unique_items = set(transaction)
    for item in unique_items:
        item_counts[item] += 1

# 5. Toplam fiş sayısı
total_receipts = len(transactions)

# 6. Support değerlerini hesaplıyorum
support_df = pd.DataFrame({
    'Ürün': list(item_counts.keys()),
    'Fiş Sayısı': list(item_counts.values()),
    'Support': [count / total_receipts for count in item_counts.values()]
})

# 7. Support değerine göre sıralıyorum büyükten küçüğe
support_df = support_df.sort_values(by='Support', ascending=False)
min_support = 0.03

# 10. Support değeri min_support'tan büyük veya eşit olanları filtrele
filtered_support_df = support_df[support_df['Support'] >= min_support]

# 11. Filtrelenmiş sonucu yeni Excel dosyasına kaydet
filtered_support_df.to_excel("C:\\Users\\edanu\\Desktop\\AprioriProject\\support_listesi_filtreli.xlsx", index=False)




print("Support listesi başarıyla oluşturuldu ve Excel'e kaydedildi.")
