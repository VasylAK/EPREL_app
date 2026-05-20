import streamlit as st
import pandas as pd
import requests
import time
import io

# 1. Zmiana tytułu strony na EPRELap_KTA
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

# --- UI STREAMLIT ---

# 1. Zmiana tytułu aplikacji
st.title("⚡ EPRELap_KTA")

# 2. Opis działania aplikacji (wymóg informacyjny o EPREL/EAN)
st.markdown("""
Aplikacja służy do odpytywania bazy EPREL o informacje o produkcie na podstawie identyfikatora, 
którym jest **Kod EPREL** lub **EAN**.
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
    
    # 3. Okno z samplem danych (podgląd wierszy i kolumn)
    st.subheader("📋 Podgląd danych (Sample)")
    st.markdown("Poniżej znajduje się skrócony widok kolumn i przykładowych danych z Twojego pliku:")
    st.dataframe(df_in.head(5), use_container_width=True)
    
    st.divider()
    
    # 4. Wskazanie kolumny / kolumn
    st.subheader("⚙️ Konfiguracja odpytywania bazy")
    st.markdown("""
    **Krok 1:** Wskaż kolumnę lub kolumny, które mają być wykorzystane do odpytania w EPREL. 
    Należy wskazać kolumny, w których znajduje się kod **EAN** lub **EPREL**.
    _Uwaga: Kolejność wyboru kolumn określa ich priorytet (pierwsza wybrana ma najwyższy priorytet)._
    """)
    
    selected_cols = st.multiselect(
        "Wybierz kolumnę/kolumny z pliku bazowego:",
        options=df_in.columns,
        help="Wymagane jest wskazanie co najmniej jednej kolumny."
    )
    
    # 5. Określenie na podstawie jakiego identyfikatora odpytywać bazę
    st.markdown("**Krok 2:** Wybierz, na podstawie jakiego identyfikatora ma być odpytywana baza (sposób budowania zapytania):")
    selected_idents = st.multiselect(
        "Wybierz typ identyfikatora / zapytania:",
        options=["Kod EPREL", "EAN"],
        default=["Kod EPREL"],
        help="Jeśli wybierzesz oba, aplikacja w razie niepowodzenia pierwszego zapytania automatycznie sprawdzi drugi identyfikator."
    )
    
    st.divider()
    
    # Uruchomienie procesu
    if st.button("Uruchom przetwarzanie danych"):
        # 4a. Warunek konieczności wybrania min. jednej kolumny (oraz identyfikatora)
        if not selected_cols:
            st.error("Błąd: Musisz wybrać minimum jedną kolumnę, aby rozpocząć!")
        elif not selected_idents:
            st.error("Błąd: Musisz wybrać minimum jeden typ identyfikatora!")
        else:
            # Listy na dane, które zostaną dopisane na końcu pliku bazowego
            models = []
            real_ids = []
            groups = []
            classes = []
            links = []
            
            # Liczniki do raportu końcowego (Punkt 7)
            unique_checked_values = set()
            unique_success_values = set()
            total_rows_success = 0
            
            progress_bar = st.progress(0)
            total_rows = len(df_in)
            
            # Główna pętla po wierszach pliku Excel
            for i, row in df_in.iterrows():
                found_data = None
                matched_value = None
                
                # 4b. Iteracja po wybranych kolumnach (Logika priorytetów kolumn)
                for col in selected_cols:
                    val = str(row[col]).split('.')[0].strip() if pd.notnull(row[col]) else ""
                    if not val or val.lower() == 'nan':
                        continue
                    
                    unique_checked_values.add(val)
                    
                    # 5. Iteracja po wybranych identyfikatorach dla danej wartości (Logika priorytetów zapytań)
                    for ident in selected_idents:
                        if ident == "Kod EPREL":
                            found_data = get_eprel_data_by_id(val, API_KEY)
                        elif ident == "EAN":
                            found_data = get_eprel_data_by_ean(val, API_KEY)
                        
                        if found_data:
                            matched_value = val
                            break  # Znaleziono dane -> przerywamy sprawdzanie kolejnych identyfikatorów dla tej wartości
                    
                    if found_data:
                        unique_success_values.add(matched_value)
                        break  # Znaleziono dane dla tej kolumny -> przerywamy sprawdzanie kolejnych kolumn (Punkt 4b spełniony)
                
                # Przetwarzanie i mapowanie pobranego JSONa
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
                    # Jeśli nie odnaleziono danych na podstawie żadnej z kolumn ani identyfikatorów
                    models.append("Brak danych")
                    real_ids.append("Brak danych")
                    groups.append("Brak danych")
                    classes.append("Brak danych")
                    links.append("Brak danych")
                
                progress_bar.progress((i + 1) / total_rows)
                time.sleep(0.05)  # Delay dla stabilności API
            
            # 6. Dane są zapisywane w nowych kolumnach pliku bazowego (Kopia df_in + nowe kolumny)
            df_out = df_in.copy()
            df_out["EPREL_Model"] = models
            df_out["EPREL_ID"] = real_ids
            df_out["EPREL_Grupa"] = groups
            df_out["EPREL_Klasa"] = classes
            df_out["EPREL_Link_PDF"] = links
            
            st.success("Przetwarzanie zakończone pomyślnie!")
            
            # 7. Krótki raport z działania aplikacji
            st.subheader("📊 Raport z przetwarzania")
            col_rep1, col_rep2, col_rep3 = st.columns(3)
            with col_rep1:
                st.metric("Sprawdzone unikalne identyfikatory", len(unique_checked_values))
            with col_rep2:
                st.metric("Identyfikatory, dla których pobrano dane", len(unique_success_values))
            with col_rep3:
                st.metric("Łączna liczba wierszy z danymi", total_rows_success)
                
            # Wyświetlenie zaktualizowanego pliku
            st.subheader("📥 Podgląd zaktualizowanego pliku bazowego")
            st.dataframe(
                df_out,
                column_config={"EPREL_Link_PDF": st.column_config.LinkColumn("Link PDF")},
                use_container_width=True
            )
            
            # Przygotowanie Excela do pobrania
            buf_excel = io.BytesIO()
            with pd.ExcelWriter(buf_excel, engine='xlsxwriter') as writer:
                df_out.to_excel(writer, index=False)
            
            st.download_button(
                label="📥 Pobierz zaktualizowany raport Excel (XLSX)",
                data=buf_excel.getvalue(),
                file_name="eprel_raport_zaktualizowany.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
