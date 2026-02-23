# Google Maps İşletme Scraper (API yok) + Excel

Bu proje Google Maps verisini **API kullanmadan web scraping** ile toplar ve Excel'e aktarır.

## Özellikler
- Ülke + şehir odaklı arama
- Otel veya istediğiniz kategoriye göre toplama
- Serbest sorgu desteği (`--query`)
- TXT dosyasından çoklu keyword yükleme (`--keywords-file`)
- Exclude kategori/terim desteği (`--exclude-categories`)
- Websitesi olmayan işletmeleri tespit etme (`has_website=False`)
- Excel çıktısında:
  - `all_businesses`
  - `no_website`
- Excel auto-filter aktif (özellikle `has_website` kolonu için)
- **Yeni:** Streamlit UI (`ui.py`) ile Chrome sekmesinden arama başlatma ve Excel indirme

## Kurulum
```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python -m playwright install chromium
```

## Streamlit UI (önerilen kullanım)
Uygulamayı tarayıcıdan çalıştırmak için:

```bash
streamlit run ui.py
```

UI davranışı:
- Country dropdown: `pycountry` ile insan okunur ülke adları (arka planda ISO2 kod)
- City dropdown: seçilen ülkeye bağlı şehir listesi (`geonamescache` offline dataset)
- Varsayılanlar: Country=`Turkey (TR)`, City=`Istanbul`
- `Start` ile scraper çalışır, sonuçlar tabloda gösterilir
- `Download Excel` ile dosya diske yazmadan (`BytesIO`) indirilir

## CLI Örnek Kullanım
### 1) İstanbul otelleri
```bash
python app.py \
  --country "Turkey" \
  --city "Istanbul" \
  --category "hotel" \
  --max-results-per-query 120 \
  --output output/istanbul_hotels.xlsx
```

### 2) Kategori + query + exclude
```bash
python app.py \
  --country "Turkey" \
  --city "Izmir" \
  --category "restaurant" \
  --query "sea view restaurant" \
  --exclude-categories "bar,night_club" \
  --output output/izmir_food.xlsx
```

### 3) Keyword TXT ile çoklu tarama
`keywords.txt`:
```txt
hotel
spa hotel
boutique hotel
pet friendly hotel
```

```bash
python app.py \
  --country "Turkey" \
  --city "Antalya" \
  --keywords-file keywords.txt \
  --exclude-categories "hostel,motel" \
  --output output/antalya_hotels.xlsx
```

## CLI Parametreler
- `--country` (zorunlu)
- `--city` (zorunlu)
- `--category` (opsiyonel)
- `--query` (opsiyonel)
- `--keywords-file` (opsiyonel)
- `--exclude-categories` (opsiyonel, virgülle ayrılmış)
- `--max-results-per-query` (varsayılan: 60)
- `--output` (varsayılan: `output/businesses.xlsx`)
- `--headed` (opsiyonel, tarayıcıyı görünür açar)

## Websitesi olmayanları Excel'de filtreleme
- `all_businesses` sayfasında `has_website` filtresinden `False` seçin.
- veya direkt `no_website` sayfasını kullanın.

## Not
Google Maps arayüzü dinamik olduğu için bazı selector'ler zamanla değişebilir; gerekirse `app.py` içindeki selector listeleri güncellenmelidir.
