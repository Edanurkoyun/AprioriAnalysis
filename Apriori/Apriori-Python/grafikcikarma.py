import pandas as pd

# Veriyi elle oluşturuyorum (sen dilersen .csv ya da .txt'den okuyabilirsin)
data = [
    ['biskuvi', 'cikolata', 'gofret', 0.0171],
    ['cikolata', 'peynir', 'sut', 0.0171],
    ['cikolata', 'gofret', 'sut', 0.0129],
    ['cikolata', 'gofret', 'makarna', 0.0114],
    ['biskuvi', 'gofret', 'sut', 0.0114],
    ['gofret', 'soguk cay', 'sut', 0.01],
    ['cikolata', 'gofret', 'kek', 0.01],
    ['cikolata', 'gazli icecek', 'gofret', 0.01],
    ['cikolata', 'makarna', 'sut', 0.01],
    ['cikolata', 'kek', 'sut', 0.0086],
    ['biskuvi', 'cikolata', 'kek', 0.0086],
    ['cikolata', 'makarna', 'tavuk', 0.0086],
    ['biskuvi', 'cikolata', 'sut', 0.0086],
    ['biskuvi', 'kahve', 'sut', 0.0071],
    ['cikolata', 'kahve', 'sut', 0.0071],
    ['cikolata', 'ekmek', 'gofret', 0.0071],
    ['ekmek', 'gofret', 'kek', 0.0071],
    ['cikolata', 'ekmek', 'peynir', 0.0071],
    ['gofret', 'makarna', 'tavuk', 0.0071],
    ['gofret', 'makarna', 'sut', 0.0057],
    ['ekmek', 'gofret', 'tavuk', 0.0057],
    ['biskuvi', 'cikolata', 'kahve', 0.0057],
    ['biskuvi', 'cikolata', 'cips', 0.0057],
    ['ekmek', 'peynir', 'sut', 0.0043],
    ['gofret', 'kek', 'sut', 0.0043],
    ['cikolata', 'ekmek', 'kek', 0.0043],
    ['biskuvi', 'gofret', 'kek', 0.0043],
    ['cikolata', 'gofret', 'tavuk', 0.0043],
    ['ekmek', 'kek', 'sut', 0.0043],
    ['ekmek', 'gofret', 'sut', 0.0029],
    ['cikolata', 'ekmek', 'sut', 0.0029],
    ['cikolata', 'ekmek', 'tavuk', 0.0029],
    ['cikolata', 'cips', 'gazli icecek', 0.0029],
    ['biskuvi', 'kek', 'sut', 0.0014],
]

# DataFrame oluştur
df = pd.DataFrame(data, columns=['Urun1', 'Urun2', 'Urun3', 'Support'])

# Ürünleri birleştir (örneğin tire ile)
df['Kombinasyon'] = df['Urun1'] + ',' + df['Urun2'] + ',' + df['Urun3']

# İstediğin sütunları seç
df_final = df[['Kombinasyon', 'Support']]

# Excel dosyasına kaydet
df_final.to_excel('C:\\Users\\edanu\\Desktop\\AprioriProject\\urun_kombinasyonlari.xlsx', index=False)

print("Excel dosyası oluşturuldu.")
