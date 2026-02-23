# Google Maps İşletme Veri Toplayıcı (Excel + Websitesiz Filtre)

Bu araç, Google Places API ile işletme verisi toplar ve Excel'e aktarır.

Yeni sürümle birlikte:
- Ülke + şehir bazlı hedefleme
- Otel gibi kategori bazlı arama
- Serbest sorgu (`--query`)
- TXT dosyasından çoklu keyword yükleme (`--keywords-file`)
- İstenmeyen kategori/terimleri dışlama (`--exclude-categories`)
- Websitesi olmayan işletmeleri ayrı sayfada görme (`no_website`)
- Excel içinde kolon filtreleri (özellikle `has_website` kolonu üzerinden kolay filtreleme)

> Not: Google Maps web scraping yerine resmi API kullanılır.

## Kurulum
```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## API Key
`.env` dosyası:

```env
GOOGLE_MAPS_API_KEY=your_api_key_here
```

## Kullanım
### 1) Oteller (şehir + ülke)
```bash
python app.py \
  --country "Turkey" \
  --city "Istanbul" \
  --category "hotel" \
  --radius 7000 \
  --max-results-per-query 120 \
  --output output/istanbul_hotels.xlsx
```

### 2) Kategori + serbest query + exclude
```bash
python app.py \
  --country "Turkey" \
  --city "Izmir" \
  --category "restaurant" \
  --query "sea view restaurant" \
  --exclude-categories "bar,night_club" \
  --output output/izmir_food.xlsx
```

### 3) Keyword dosyasıyla çoklu arama
`keywords.txt` örneği:
```txt
hotel
spa hotel
boutique hotel
pet friendly hotel
```

Çalıştırma:
```bash
python app.py \
  --country "Turkey" \
  --city "Antalya" \
  --keywords-file keywords.txt \
  --exclude-categories "hostel,motel" \
  --output output/antalya_hotels.xlsx
```

## Parametreler
- `--country` (zorunlu): Ülke adı
- `--city` (zorunlu): Şehir adı
- `--category`: Kategori (hotel, cafe, dental clinic, vb.)
- `--query`: Serbest metin sorgu
- `--keywords-file`: Satır satır keyword içeren txt dosyası
- `--exclude-categories`: Virgülle ayrılmış dışlanacak terimler
- `--location`: Opsiyonel `lat,lng` (verilirse geocode atlanır)
- `--radius`: Metre (varsayılan 5000)
- `--max-results-per-query`: Her sorgu için üst limit (varsayılan 60)
- `--output`: Excel dosya yolu

## Excel Çıktısı
- `all_businesses`: Tüm sonuçlar
- `no_website`: `has_website=False` olanlar

Her iki sayfada da Excel filtreleri aktiftir. `all_businesses` sayfasında `has_website` filtresinden sadece `False` seçerek websitesiz işletmeleri anında filtreleyebilirsiniz.
