import argparse
import os
import time
from pathlib import Path
from typing import Dict, List, Optional, Set

import pandas as pd
import requests
from dotenv import load_dotenv
from openpyxl.utils import get_column_letter

TEXT_SEARCH_URL = "https://maps.googleapis.com/maps/api/place/textsearch/json"
DETAILS_URL = "https://maps.googleapis.com/maps/api/place/details/json"
GEOCODE_URL = "https://maps.googleapis.com/maps/api/geocode/json"


def load_keywords_from_txt(path: Optional[str]) -> List[str]:
    if not path:
        return []

    file_path = Path(path)
    if not file_path.exists():
        raise FileNotFoundError(f"Keyword dosyası bulunamadı: {path}")

    keywords: List[str] = []
    for line in file_path.read_text(encoding="utf-8").splitlines():
        normalized = line.strip()
        if normalized and not normalized.startswith("#"):
            keywords.append(normalized)

    return keywords


def geocode_city_country(api_key: str, city: str, country: str) -> str:
    query = f"{city}, {country}"
    params = {"address": query, "key": api_key, "language": "tr"}
    resp = requests.get(GEOCODE_URL, params=params, timeout=30)
    resp.raise_for_status()
    payload = resp.json()

    status = payload.get("status")
    if status != "OK" or not payload.get("results"):
        raise RuntimeError(f"Geocode API hata durumu: {status} | {payload}")

    loc = payload["results"][0]["geometry"]["location"]
    return f"{loc['lat']},{loc['lng']}"


def text_search(api_key: str, query: str, location: str, radius: int, max_results: int) -> List[Dict]:
    results: List[Dict] = []
    next_page_token: Optional[str] = None

    while len(results) < max_results:
        params = {
            "key": api_key,
            "query": query,
            "location": location,
            "radius": radius,
            "language": "tr",
        }

        if next_page_token:
            params = {"key": api_key, "pagetoken": next_page_token}
            time.sleep(2)

        resp = requests.get(TEXT_SEARCH_URL, params=params, timeout=30)
        resp.raise_for_status()
        payload = resp.json()

        status = payload.get("status")
        if status not in {"OK", "ZERO_RESULTS"}:
            raise RuntimeError(f"Text Search API hata durumu: {status} | {payload}")

        page_results = payload.get("results", [])
        if not page_results:
            break

        results.extend(page_results)
        if len(results) >= max_results:
            break

        next_page_token = payload.get("next_page_token")
        if not next_page_token:
            break

    return results[:max_results]


def place_details(api_key: str, place_id: str) -> Dict:
    fields = ",".join(
        [
            "place_id",
            "name",
            "formatted_address",
            "formatted_phone_number",
            "website",
            "rating",
            "user_ratings_total",
            "types",
            "url",
            "business_status",
        ]
    )

    params = {"key": api_key, "place_id": place_id, "fields": fields, "language": "tr"}
    resp = requests.get(DETAILS_URL, params=params, timeout=30)
    resp.raise_for_status()
    payload = resp.json()

    status = payload.get("status")
    if status != "OK":
        raise RuntimeError(f"Details API hata durumu: {status} | {payload}")

    return payload.get("result", {})


def build_queries(
    city: str,
    country: str,
    category: Optional[str],
    direct_query: Optional[str],
    keywords: List[str],
) -> List[str]:
    base_queries: List[str] = []

    if direct_query:
        base_queries.append(direct_query.strip())

    if category:
        base_queries.append(category.strip())

    base_queries.extend(keywords)

    if not base_queries:
        raise ValueError("En az bir sorgu kaynağı verin: --query veya --category veya --keywords-file")

    final_queries: List[str] = []
    seen: Set[str] = set()

    for item in base_queries:
        normalized = " ".join(item.split()).strip()
        if not normalized:
            continue
        full_query = f"{normalized} in {city}, {country}"
        key = full_query.lower()
        if key not in seen:
            seen.add(key)
            final_queries.append(full_query)

    return final_queries


def should_exclude(details: Dict, exclude_terms: List[str]) -> bool:
    if not exclude_terms:
        return False

    normalized_terms = [term.lower().strip() for term in exclude_terms if term.strip()]
    if not normalized_terms:
        return False

    name = str(details.get("name", "")).lower()
    types_blob = " ".join(details.get("types", [])).lower()

    for term in normalized_terms:
        if term in name or term in types_blob:
            return True

    return False


def collect_businesses(
    api_key: str,
    queries: List[str],
    location: str,
    radius: int,
    max_results_per_query: int,
    exclude_terms: List[str],
) -> pd.DataFrame:
    rows: List[Dict] = []
    seen_place_ids: Set[str] = set()

    for query in queries:
        search_hits = text_search(api_key, query, location, radius, max_results_per_query)

        for hit in search_hits:
            place_id = hit.get("place_id")
            if not place_id or place_id in seen_place_ids:
                continue

            details = place_details(api_key, place_id)
            if should_exclude(details, exclude_terms):
                continue

            seen_place_ids.add(place_id)
            website = details.get("website")

            rows.append(
                {
                    "source_query": query,
                    "place_id": details.get("place_id"),
                    "name": details.get("name"),
                    "address": details.get("formatted_address"),
                    "phone": details.get("formatted_phone_number"),
                    "website": website,
                    "has_website": bool(website and str(website).strip()),
                    "rating": details.get("rating"),
                    "user_ratings_total": details.get("user_ratings_total"),
                    "types": ", ".join(details.get("types", [])),
                    "business_status": details.get("business_status"),
                    "google_maps_url": details.get("url"),
                }
            )

    df = pd.DataFrame(rows)
    if not df.empty:
        df = df.sort_values(by=["has_website", "rating", "user_ratings_total"], ascending=[True, False, False])
    return df


def apply_excel_filters_and_widths(writer: pd.ExcelWriter, sheet_name: str, df: pd.DataFrame) -> None:
    ws = writer.sheets[sheet_name]
    if df.empty:
        return

    max_col = df.shape[1]
    max_row = df.shape[0] + 1
    ws.auto_filter.ref = f"A1:{get_column_letter(max_col)}{max_row}"

    for idx, column in enumerate(df.columns, start=1):
        max_len = max(df[column].astype(str).map(len).max(), len(column))
        ws.column_dimensions[get_column_letter(idx)].width = min(max_len + 2, 60)



def export_excel(df: pd.DataFrame, output_path: str) -> None:
    out = Path(output_path)
    out.parent.mkdir(parents=True, exist_ok=True)

    no_website_df = df[df["has_website"] == False].copy() if not df.empty else df.copy()  # noqa: E712

    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="all_businesses", index=False)
        no_website_df.to_excel(writer, sheet_name="no_website", index=False)

        apply_excel_filters_and_widths(writer, "all_businesses", df)
        apply_excel_filters_and_widths(writer, "no_website", no_website_df)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Google Maps işletme verisi toplayıp Excel'e aktarır.")

    parser.add_argument("--country", required=True, help="Ülke adı (örn: Turkey)")
    parser.add_argument("--city", required=True, help="Şehir adı (örn: Istanbul)")

    parser.add_argument("--category", help="Kategori (örn: hotel, dental clinic, restaurant)")
    parser.add_argument("--query", help="Serbest sorgu (kategori yerine veya kategoriyle birlikte kullanılabilir)")
    parser.add_argument("--keywords-file", help="Satır satır keyword içeren txt dosyası")

    parser.add_argument(
        "--exclude-categories",
        default="",
        help="Virgülle ayrılmış hariç tutulacak kategori/terimler (örn: bar,night_club,casino)",
    )

    parser.add_argument("--location", help="Opsiyonel: lat,lng. Verilirse geocode yerine bu kullanılır")
    parser.add_argument("--radius", type=int, default=5000, help="Arama yarıçapı (metre)")
    parser.add_argument("--max-results-per-query", type=int, default=60, help="Her sorgu için maksimum işletme")
    parser.add_argument("--output", default="output/businesses.xlsx", help="Excel çıktı dosyası")
    return parser.parse_args()


def main() -> None:
    load_dotenv()
    args = parse_args()

    api_key = os.getenv("GOOGLE_MAPS_API_KEY")
    if not api_key:
        raise EnvironmentError("GOOGLE_MAPS_API_KEY bulunamadı. Lütfen .env dosyanızı kontrol edin.")

    keywords = load_keywords_from_txt(args.keywords_file)
    exclude_terms = [item.strip() for item in args.exclude_categories.split(",") if item.strip()]

    location = args.location or geocode_city_country(api_key, args.city, args.country)
    queries = build_queries(
        city=args.city,
        country=args.country,
        category=args.category,
        direct_query=args.query,
        keywords=keywords,
    )

    df = collect_businesses(
        api_key=api_key,
        queries=queries,
        location=location,
        radius=args.radius,
        max_results_per_query=args.max_results_per_query,
        exclude_terms=exclude_terms,
    )

    export_excel(df, args.output)
    total = len(df)
    no_website_count = int((~df["has_website"]).sum()) if total else 0

    print(f"Konum: {args.city}, {args.country} | center={location}")
    print(f"Sorgu sayısı: {len(queries)}")
    print(f"Toplam işletme: {total}")
    print(f"Websitesi olmayan işletme: {no_website_count}")
    print(f"Excel yazıldı: {args.output}")


if __name__ == "__main__":
    main()
