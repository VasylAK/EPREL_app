import streamlit as st
import pandas as pd
import requests
import time
import io

# --- KONFIGURACJA STRONY ---
st.set_page_config(page_title="EPRELap_KTA", page_icon="⚡", layout="wide")

# --- FUNKCJE POMOCNICZE DO API ---

def get_eprel_data_by_id(eprel_id, api_key):
    """Odpytanie bazy przy użyciu Kodu EPREL."""
    url = f"https://eprel.ec.europa.eu/api/product/{eprel_id.strip()}"
    headers = {"Authorization": f"Bearer {api_key}", "Accept": "application/json"}
    try:
        response = requests.get(url, headers=headers, timeout=12)
        if response.status_code == 200:
            return response.json()
    except Exception:
        pass
    return None

def get_eprel_data_by_ean(ean, api_key):
    """Odpytanie bazy przy użyciu kodu EAN (GTIN)."""
    url = f"https://eprel.ec.europa.eu/api/product/gtin/{ean.strip()}"
    headers = {"Authorization": f"Bearer {api_key}", "Accept": "application/json"}
    try:
        response = requests.get(url, headers=headers, timeout=12)
        if response.status_code == 200:
            return response.json()
    except Exception:
        pass
    return None

# --- FUNKCJE WALIDACYJNE ---

def is_valid_eprel(val):
    """Kod EPREL powinien składać się z samych cyfr."""
    return val.isdigit()

def is_valid_ean(val):
    """Kod EAN (GTIN) powinien składać się z samych cyfr i mieć odpowiednią długość."""
    return val.isdigit() and len(val) in [8, 12, 13, 14]

# --- UI STREAMLIT ---

st.title("⚡ EPRELap_KTA")
st.markdown("""
Aplikacja służy do odpytywania bazy EPREL o informacje o produkcie na podstawie identyfikatora, 
którym jest **Kod EPREL** lub **EAN**. Zoptymalizowana wersja korzysta z walidacji oraz pamięci podręcznej.
""")

# Pobieranie klucza z Secrets
try:
    API_KEY = st.secrets["EPREL_API_KEY"]
except Exception:
    st.error("Błąd: Nie znaleziono klucza 'EPREL_API_KEY' w Secrets na Streamlit Cloud!")
    st.stop()

# Wczytywanie pliku
uploaded_file = st.file_uploader("Załaduj plik Excel (.xlsx)", type=["xlsx"])

if uploaded_file:
    df_in = pd.read_excel(uploaded_file)
    
    st.subheader("📋 Podgląd danych (Sample)")
    st.markdown("Poniżej znajduje się skrócony widok kolumn i przykładowych danych z Twojego pliku:")
    st.dataframe(df_in.head(5), use_container_width=True)
    
    st.divider()
    
    st.subheader("⚙️ Konfiguracja odpytywania bazy")
    st.markdown("""
    **Krok 1:** Wskaż kolumnę lub kolumny, które mają być wykorzystane do odpytania w EPREL. 
    Należy wskazać kolumny, w których znajduje się kod **EAN** lub **EPREL**.
    _Uwaga: Kolejność wyboru kolumn określa ich priorytet._
    """)
    
    selected_cols = st.multiselect(
        "Wybierz kolumnę/kolumny z pliku bazowego:",
        options=df_in.columns,
        help="Wymagane jest wskazanie co najmniej jednej kolumny."
    )
    
    st.markdown("**Krok 2:** Wybierz, na podstawie jakiego identyfikatora ma być odpytywana baza:")
    selected_idents = st.multiselect(
        "Wybierz typ identyfikatora / zapytania:",
        options=["Kod EPREL", "EAN"],
        default=["Kod EPREL"],
        help="Jeśli wybierzesz oba, aplikacja wykona zapytania kaskadowo dla poprawnego formatu."
    )
    
    st.divider()
    
    # Uruchomienie procesu
    if st.button("Uruchom przetwarzanie danych"):
        if not selected_cols:
            st.error("Błąd: Musisz wybrać minimum jedną kolumnę, aby rozpocząć!")
        elif not selected_idents:
            st.error("Błąd: Musisz wybrać minimum jeden typ identyfikatora!")
        else:
            models, real_ids, groups, classes, links = [], [], [], [], []
            
            # --- PAMIĘĆ PODRĘCZNA (CACHE) ---
            # Słowniki zapobiegające powtórnym zapytaniom dla tych samych kodów
            cache_eprel = {}
            cache_ean = {}
            
            unique_checked_values = set()
            unique_success_values = set()
            total_rows_success = 0
            
            progress_bar = st.progress(0)
            total_rows = len(df_in)
            
            for i, row in df_in.iterrows():
                found_data = None
                matched_value = None
                
                # Iteracja po wybranych kolumnach
                for col in selected_cols:
                    # Czyszczenie wartości - usunięcie .0 jeśli pandas potraktował cyfry jako float
                    val = str(row[col]).split('.')[0].strip() if pd.notnull(row[col]) else ""
                    if not val or val.lower() == 'nan':
                        continue
                    
                    unique_checked_values.add(val)
                    
                    # Iteracja po identyfikatorach z WALIDACJĄ I CACHE
                    for ident in selected_idents:
                        
                        if ident == "Kod EPREL" and is_valid_eprel(val):
                            # Jeśli nie ma w pamięci, zapytaj API i zapisz
                            if val not in cache_eprel:
                                cache_eprel[val] = get_eprel_data_by_id(val, API_KEY)
                                time.sleep(0.05) # Opóźnienie tylko po realnym strzale do API
                            
                            found_data = cache_eprel[val]
                            
                        elif ident == "EAN" and is_valid_ean(val):
                            if val not in cache_ean:
                                cache_ean[val] = get_eprel_data_by_ean(val, API_KEY)
                                time.sleep(0.05)
                            
                            found_data = cache_ean[val]
                        
                        if found_data:
                            matched_value = val
                            break # Przerywamy identyfikatory
                    
                    if found_data:
                        unique_success_values.add(matched_value)
                        break # Przerywamy kolumny
                
                # Zapis wyników
                if found_data:
                    total_rows_success += 1
                    energy_class = found_data.get("energyClass", "N/A")
                    model_identifier = found_data.get("modelIdentifier", "N/A")
                    real_id = found_data.get("registrationNumber") or found_data.get("eprelRegistrationNumber") or ""
                    product_group = found_data.get("productGroup", "N/A")
                    
                    if product_group != "N/A" and real_id:
                        pdf_link = f"https://eprel.ec.europa.eu/labels/{product_group}/Label_{real_id}.pdf"
                    else:
                        pdf_link = "Brak grupy produktowej"
                        
                    models.append(model_identifier)
                    real_ids.append(real_id)
                    groups.append(product_group)
                    classes.append(energy_class)
                    links.append(pdf_link)
                else:
                    models.append("Brak danych")
                    real_ids.append("Brak danych")
                    groups.append("Brak danych")
                    classes.append("Brak danych")
                    links.append("Brak danych")
                
                progress_bar.progress((i + 1) / total_rows)
            
            df_out = df_in.copy()
            df_out["EPREL_Model"] = models
            df_out["EPREL_ID"] = real_ids
            df_out["EPREL_Grupa"] = groups
            df_out["EPREL_Klasa"] = classes
            df_out["EPREL_Link_PDF"] = links
            
            st.success("Przetwarzanie zakończone pomyślnie!")
            
            st.subheader("📊 Raport z przetwarzania")
            col_rep1, col_rep2, col_rep3, col_rep4 = st.columns(4)
            with col_rep1:
                st.metric("Sprawdzone identyfikatory", len(unique_checked_values))
            with col_rep2:
                st.metric("Sukces (ilość unikalnych)", len(unique_success_values))
            with col_rep3:
                st.metric("Uzupełnione wiersze", total_rows_success)
            with col_rep4:
                # Informacja z cache
                api_calls = len(cache_eprel) + len(cache_ean)
                st.metric("Rzeczywiste zapytania API", api_calls, help="Ilość faktycznych zapytań do serwera (dzięki cache to mniej niż liczba wierszy).")
                
            st.subheader("📥 Podgląd zaktualizowanego pliku")
            st.dataframe(
                df_out,
                column_config={"EPREL_Link_PDF": st.column_config.LinkColumn("Link PDF")},
                use_container_width=True
            )
            
            buf_excel = io.BytesIO()
            with pd.ExcelWriter(buf_excel, engine='xlsxwriter') as writer:
                df_out.to_excel(writer, index=False)
            
            st.download_button(
                label="📥 Pobierz zaktualizowany raport Excel (XLSX)",
                data=buf_excel.getvalue(),
                file_name="eprel_raport_zaktualizowany.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
