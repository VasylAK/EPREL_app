import streamlit as st
import pandas as pd
import requests
import time
import io

# --- KONFIGURACJA STRONY ---
st.set_page_config(page_title="EPREL Link Generator", page_icon="⚡", layout="wide")

# --- FUNKCJE POMOCNICZE ---

def get_eprel_json(eprel_id, api_key):
    """
    KROK 1: Pobiera metadane (JSON), abyśmy znali 'productGroup'
    niezbędne do zbudowania linku.
    """
    if not eprel_id or str(eprel_id).lower() == 'nan' or str(eprel_id).strip() == "":
        return None

    # Endpoint zwracający szczegóły produktu
    url = f"https://eprel.ec.europa.eu/api/product/{eprel_id.strip()}"

    headers = {
        "Authorization": f"Bearer {api_key}",
        "Accept": "application/json"
    }
    
    try:
        response = requests.get(url, headers=headers, timeout=10)
        if response.status_code == 200:
            return response.json()
        return None
    except Exception:
        return None

# --- UI STREAMLIT ---
st.title("⚡ EPREL: Generator Linków do Etykiet")
st.markdown("""
Aplikacja pobiera dane z API i generuje bezpośrednie linki do etykiet PDF w pliku Excel.
**Nie pobiera fizycznych plików.**
""")

# Pobieranie klucza z Secrets
try:
    API_KEY = st.secrets["EPREL_API_KEY"]
except Exception:
    st.error("Błąd: Nie znaleziono klucza 'EPREL_API_KEY' w Secrets!")
    st.stop()

uploaded_file = st.file_uploader("Załaduj plik Excel (wymagana kolumna: 'kod eprel')", type=["xlsx"])

if uploaded_file:
    df_in = pd.read_excel(uploaded_file)
    cols_lower = [str(c).lower() for c in df_in.columns]
    
    # Sprawdzenie czy istnieje wymagana kolumna
    if 'kod eprel' not in cols_lower:
        st.error("Plik musi zawierać kolumnę: 'kod eprel'")
    else:
        if st.button("Generuj raport z linkami"):
            final_data = []
            progress_bar = st.progress(0)
            
            # Znalezienie właściwej nazwy kolumny (bez względu na wielkość liter)
            code_col = [c for c in df_in.columns if c.lower() == 'kod eprel'][0]
            
            total_rows = len(df_in)

            for i, row in df_in.iterrows():
                # Przygotowanie ID z pliku wejściowego
                eprel_id_input = str(row[code_col]).split('.')[0].strip() if pd.notnull(row[code_col]) else ""
                
                # Domyślny wpis (w razie błędu API)
                entry = {
                    "Kod EPREL (Input)": eprel_id_input,
                    "Model": "Brak danych",
                    "Kod EPREL (API)": "",
                    "Grupa Produktowa": "",
                    "Klasa Energetyczna": "",
                    "Link do Etykiety (PDF)": "Błąd / Nie znaleziono"
                }

                if eprel_id_input:
                    # --- KROK 1: Pobranie JSON z API ---
                    data = get_eprel_json(eprel_id_input, API_KEY)
                    
                    if data:
                        # --- KROK 2: Mapowanie danych ---
                        energy_class = data.get("energyClass", "N/A")
                        model_identifier = data.get("modelIdentifier", "N/A")
                        # API może zwrócić inny numer rejestracyjny niż podany (np. przekierowanie)
                        real_id = data.get("registrationNumber") or data.get("eprelRegistrationNumber") or eprel_id_input
                        product_group = data.get("productGroup", "N/A")

                        entry["Model"] = model_identifier
                        entry["Kod EPREL (API)"] = real_id
                        entry["Grupa Produktowa"] = product_group
                        entry["Klasa Energetyczna"] = energy_class

                        # --- KROK 3: Budowa linku (Tekst) ---
                        if product_group != "N/A":
                            # Wzór: .../labels/[grupa]/Label_[id].pdf
                            pdf_link = f"https://eprel.ec.europa.eu/labels/{product_group}/Label_{real_id}.pdf"
                            entry["Link do Etykiety (PDF)"] = pdf_link
                        else:
                            entry["Link do Etykiety (PDF)"] = "Brak Grupy Produktowej"

                final_data.append(entry)
                
                # Aktualizacja paska postępu
                progress_bar.progress((i + 1) / total_rows)
                time.sleep(0.05) # Delay dla API

            # Zapis wyników do sesji
            st.session_state.results_df = pd.DataFrame(final_data)
            st.success("Zakończono! Linki zostały wygenerowane.")

# --- WYŚWIETLANIE I POBIERANIE ---
if 'results_df' in st.session_state:
    st.subheader("Podgląd wyników")
    
    # Konfiguracja wyświetlania tabeli - kolumna z linkiem jest klikalna
    st.dataframe(
        st.session_state.results_df,
        column_config={
            "Link do Etykiety (PDF)": st.column_config.LinkColumn("Link PDF")
        },
        use_container_width=True
    )
    
    # Generowanie pliku Excel w pamięci
    buf_excel = io.BytesIO()
    with pd.ExcelWriter(buf_excel, engine='xlsxwriter') as writer:
        st.session_state.results_df.to_excel(writer, index=False)
    
    # Przycisk pobierania
    st.download_button(
        label="📥 Pobierz raport Excel (XLSX)",
        data=buf_excel.getvalue(),
        file_name="eprel_linki_etykiet.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
