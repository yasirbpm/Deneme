import io
from datetime import datetime
from typing import Dict, List, Tuple

import geonamescache
import pandas as pd
import pycountry
import streamlit as st

from app import build_queries, collect_businesses


def sanitize_for_filename(value: str) -> str:
    cleaned = "".join(ch if ch.isalnum() or ch in {"-", "_"} else "_" for ch in value.strip())
    return "_".join(part for part in cleaned.split("_") if part) or "output"


@st.cache_data
def load_countries() -> List[Tuple[str, str]]:
    countries = []
    for country in pycountry.countries:
        alpha_2 = getattr(country, "alpha_2", None)
        if alpha_2:
            countries.append((country.name, alpha_2))

    countries.sort(key=lambda item: item[0])
    return countries


@st.cache_data
def load_cities_by_country() -> Dict[str, List[str]]:
    gc = geonamescache.GeonamesCache()
    cities = gc.get_cities()

    mapping: Dict[str, set] = {}
    for city in cities.values():
        countrycode = city.get("countrycode")
        name = city.get("name")
        if countrycode and name:
            mapping.setdefault(countrycode, set()).add(name)

    return {country: sorted(list(names)) for country, names in mapping.items()}


@st.cache_data
def country_name_to_code_map() -> Dict[str, str]:
    return {name: code for name, code in load_countries()}


@st.cache_data
def country_code_to_name_map() -> Dict[str, str]:
    return {code: name for name, code in load_countries()}


def build_excel_bytes(df: pd.DataFrame) -> bytes:
    no_website_df = df[df["has_website"] == False].copy() if not df.empty else df.copy()  # noqa: E712

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="all_businesses", index=False)
        no_website_df.to_excel(writer, sheet_name="no_website", index=False)

    output.seek(0)
    return output.read()


def main() -> None:
    st.set_page_config(page_title="Google Maps Scraper UI", layout="wide")
    st.title("Google Maps İşletme Scraper")
    st.caption("CLI yerine tarayıcıdan arama başlatın ve Excel indirin.")

    countries = load_countries()
    cities_by_country = load_cities_by_country()
    name_to_code = country_name_to_code_map()
    code_to_name = country_code_to_name_map()

    country_names = [name for name, _ in countries]
    default_country_name = code_to_name.get("TR", "Turkey")
    default_country_index = country_names.index(default_country_name) if default_country_name in country_names else 0

    country_name = st.selectbox("Country", options=country_names, index=default_country_index)
    country_code = name_to_code[country_name]

    city_options = cities_by_country.get(country_code, [])
    if not city_options:
        st.warning("Bu ülke için şehir listesi bulunamadı. Lütfen farklı ülke seçin.")
        st.stop()

    default_city = "Istanbul" if country_code == "TR" and "Istanbul" in city_options else city_options[0]
    default_city_index = city_options.index(default_city)
    city_name = st.selectbox("City", options=city_options, index=default_city_index)

    category = st.text_input("Category (optional)", placeholder="hotel / restaurant / dental clinic")
    query = st.text_input("Query / Keyword (optional)", placeholder="sea view hotel")
    exclude_categories = st.text_input("Exclude categories (comma-separated, optional)", placeholder="bar,night_club")

    max_results_per_query = st.number_input("Max results per query", min_value=1, max_value=1000, value=100, step=1)
    headed = st.checkbox("Headed (show browser during scraping)", value=True)

    start = st.button("Start", type="primary")

    if start:
        if not category.strip() and not query.strip():
            st.error("En az bir giriş gerekli: Category veya Query / Keyword.")
            st.stop()

        exclude_terms = [item.strip() for item in exclude_categories.split(",") if item.strip()]

        with st.status("Scraping başlatılıyor...", expanded=True) as status:
            st.write("Sorgular hazırlanıyor...")
            queries = build_queries(
                city=city_name,
                country=country_name,
                category=category.strip() or None,
                direct_query=query.strip() or None,
                keywords=[],
            )
            st.write(f"Toplam sorgu: {len(queries)}")
            st.write("Google Maps taraması çalışıyor...")

            df = collect_businesses(
                queries=queries,
                max_results_per_query=int(max_results_per_query),
                exclude_terms=exclude_terms,
                headless=not headed,
            )

            status.update(label="Tamamlandı", state="complete")

        st.success(f"Toplam işletme: {len(df)}")
        if not df.empty:
            no_website_count = int((~df["has_website"]).sum())
            st.info(f"Websitesi olmayan işletme: {no_website_count}")
        else:
            st.warning("Sonuç bulunamadı.")

        st.dataframe(df, use_container_width=True)

        excel_bytes = build_excel_bytes(df)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M")
        filename = f"{sanitize_for_filename(country_name)}_{sanitize_for_filename(city_name)}_{timestamp}.xlsx"

        st.download_button(
            label="Download Excel",
            data=excel_bytes,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


if __name__ == "__main__":
    main()
