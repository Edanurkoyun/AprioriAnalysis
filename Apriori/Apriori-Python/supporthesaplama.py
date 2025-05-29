import pandas as pd
from collections import defaultdict

# 1. Excel dosyasını oku
df = pd.read_excel("C:\\Users\\edanu\\Desktop\\AprioriProject\\temizlenmis_veri2.xlsx", header=None)

# 2. İlk sütunu çıkar (Fiş numarası)
urun_df = df.iloc[:, 1:]

# 3. Her satırı fiş olarak işle, verileri temizle: NaN -> string -> strip -> lower
transactions = urun_df.apply(lambda row: [str(item).strip().lower() for item in row.dropna()], axis=1).tolist()

# 4. Ürünleri say (aynı fişte tekrar eden ürünleri bir kez say)
item_counts = defaultdict(int)
for transaction in transactions:
    unique_items = set(transaction)
    for item in unique_items:
        item_counts[item] += 1

# 5. Toplam fiş sayısı
total_receipts = len(transactions)

# 6. Support değerlerini hesapla
support_df = pd.DataFrame({
    'Ürün': list(item_counts.keys()),
    'Fiş Sayısı': list(item_counts.values()),
    'Support': [count / total_receipts for count in item_counts.values()]
})

# 7. Support değerine göre sırala
support_df = support_df.sort_values(by='Support', ascending=False)

# 8. Excel'e kaydet
support_df.to_excel("C:\\Users\\edanu\\Desktop\\AprioriProject\\support_listesi.xlsx", index=False)

print("Support listesi başarıyla oluşturuldu ve Excel'e kaydedildi.")
