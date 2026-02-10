import streamlit as st
import pandas as pd
import requests
import time
import io
import zipfile

# --- KONFIGURACJA STRONY ---
st.set_page_config(page_title="EPREL Data & File Downloader", page_icon="⚡", layout="wide")

# --- FUNKCJE POMOCNICZE ---

def get_eprel_data(eprel_id, ean, api_key):
    """Pobiera dane produktu z API EPREL."""
    if eprel_id and str(eprel_id).lower() != 'nan' and str(eprel_id).strip() != "":
        url = f"https://eprel.ec.europa.eu/api/product/{eprel_id.strip()}"
    elif ean and str(ean).lower() != 'nan' and str(ean).strip() != "":
        url = f"https://eprel.ec.europa.eu/api/product/gtin/{ean.strip()}"
    else:
        return None

    headers = {
        "Authorization": f"Bearer {api_key}",
        "Accept": "application/json"
    }
    
    try:
        response = requests.get(url, headers=headers, timeout=15)
        if response.status_code == 200:
            return response.json()
        return None
    except Exception:
        return None

def download_eprel_file(url, api_key):
    """Pobiera plik binarny z API EPREL przy użyciu autoryzacji."""
    headers = {"Authorization": f"Bearer {api_key}"}
    try:
        response = requests.get(url, headers=headers, timeout=20)
        if response.status_code == 200:
            return response.content
        return None
    except Exception:
        return None

# --- UI STREAMLIT ---
st.title("⚡ EPREL Pro: Dane, Linki i Załączniki ZIP")
st.markdown("""
Aplikacja pobiera klasę energetyczną, generuje publiczne linki do produktów oraz przygotowuje paczkę ZIP 
z etykietami (PNG) i kartami (PDF) nazwanymi według Twojego numeru EAN.
""")

# Pobieranie klucza z Secrets
try:
    API_KEY = st.secrets["EPREL_API_KEY"]
except Exception:
    st.error("Błąd: Nie znaleziono klucza 'EPREL_API_KEY' w Secrets na Streamlit Cloud!")
    st.stop()

uploaded_file = st.file_uploader("Załaduj plik Excel (wymagane kolumny: 'ean' oraz 'kod eprel')", type=["xlsx"])

if uploaded_file:
    df_in = pd.read_excel(uploaded_file)
    
    # Normalizacja nazw kolumn do małych liter dla łatwiejszego wyszukiwania
    cols_lower = [str(c).lower() for c in df_in.columns]
    
    if 'ean' not in cols_lower or 'kod eprel' not in cols_lower:
        st.error("Plik musi zawierać kolumny: 'ean' i 'kod eprel'")
    else:
        if st.button("Uruchom przetwarzanie"):
            final_data = []
            zip_buffer = io.BytesIO()
            progress_bar = st.progress(0)
            
            # Znalezienie oryginalnych nazw kolumn
            ean_col = [c for c in df_in.columns if c.lower() == 'ean'][0]
            code_col = [c for c in df_in.columns if c.lower() == 'kod eprel'][0]

            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                for i, row in df_in.iterrows():
                    # Przygotowanie danych wejściowych
                    ean_val = str(row[ean_col]).split('.')[0].strip() if pd.notnull(row[ean_col]) else f"brak_ean_{i}"
                    eprel_id_val = str(row[code_col]).split('.')[0].strip() if pd.notnull(row[code_col]) else ""
                    
                    entry = {
                        "EAN": ean_val,
                        "Kod EPREL (Input)": eprel_id_val,
                        "Klasa Energetyczna": "Nie znaleziono",
                        "Link do produktu": "Brak",
                        "Status plików": "Błąd"
                    }

                    # Pobieranie danych z API
                    data = get_eprel_data(eprel_id_val, ean_val, API_KEY)
                    
                    if data:
                        # Wyciągnięcie realnego ID rejestracyjnego z bazy
                        real_id = data.get("registrationNumber") or eprel_id_val
                        entry["Klasa Energetyczna"] = data.get("energyClass", "N/A")
                        
                        # --- TRANSFORMACJA LINKU ---
                        # Generujemy "wyczyszczony" publiczny link do produktu (zamiast technicznego linku API)
                        entry["Link do produktu"] = f"https://eprel.ec.europa.eu/screen/product/productModel/{real_id}"
                        
                        # --- POBIERANIE PLIKÓW DO PACZKI ZIP ---
                        # 1. Etykieta (PNG)
                        label_url = f"https://eprel.ec.europa.eu/api/product/{real_id}/label?format=PNG"
                        label_bits = download_eprel_file(label_url, API_KEY)
                        if label_bits:
                            zip_file.writestr(f"etykiety/{ean_val}.png", label_bits)
                        
                        # 2. Karta produktu (PDF)
                        fiche_url = f"https://eprel.ec.europa.eu/api/product/{real_id}/fiches"
                        fiche_bits = download_eprel_file(fiche_url, API_KEY)
                        if fiche_bits:
                            zip_file.writestr(f"karty/{ean_val}.pdf", fiche_bits)
                        
                        entry["Status plików"] = "Pobrano"

                    final_data.append(entry)
                    progress_bar.progress((i + 1) / len(df_in))
                    time.sleep(0.05) # Delikatny delay dla stabilności API

            # Zapis wyników do sesji
            st.session_state.results_df = pd.DataFrame(final_data)
            st.session_state.zip_data = zip_buffer.getvalue()
            st.success("Przetwarzanie zakończone!")

# --- WYŚWIETLANIE WYNIKÓW I POBIERANIE ---
if 'results_df' in st.session_state:
    st.subheader("Podgląd danych")
    st.dataframe(
        st.session_state.results_df, 
        column_config={"Link do produktu": st.column_config.LinkColumn("Otwórz w EPREL")}
    )
    
    col1, col2 = st.columns(2)
    
    with col1:
        # Generowanie pliku Excel do pobrania
        buf_excel = io.BytesIO()
        with pd.ExcelWriter(buf_excel, engine='xlsxwriter') as writer:
            st.session_state.results_df.to_excel(writer, index=False)
        
        st.download_button(
            label="📥 Pobierz raport Excel",
            data=buf_excel.getvalue(),
            file_name="eprel_raport_wynikowy.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    
    with col2:
        # Pobieranie gotowej paczki ZIP
        st.download_button(
            label="📦 Pobierz paczkę ZIP (Etykiety i Karty)",
            data=st.session_state.zip_data,
            file_name="zalaczniki_eprel_ean.zip",
            mime="application/zip"
        )
