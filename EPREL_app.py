import streamlit as st
import pandas as pd
import requests
import time
import io
import concurrent.futures
import threading
import random

# --- 1. CONFIGURATION ---
st.set_page_config(page_title="EPRELap_KTA - Async Edition", page_icon="⚡", layout="wide")

# Blokada do bezpiecznej obsługi Cache w wielu wątkach jednocześnie
cache_lock = threading.Lock()
cache_eprel = {}
cache_ean = {}

# --- ZAAWANSOWANA FUNKCJA API Z EXPONENTIAL BACKOFF (ODPORNOŚĆ NA 429) ---

def call_eprel_api_with_retry(url, headers, session):
    """
    Wykonuje zapytanie HTTP z zaawansowaną logiką ponawiania prób.
    W przypadku błędu 429 aplikacja usypia wątek i próbuje ponownie później.
    """
    max_retries = 5  # Maksymalnie 5 prób dla jednej pozycji w przypadku przeciążenia
    base_delay = 2   # Bazowe opóźnienie w sekundach
    
    for attempt in range(max_retries):
        try:
            response = session.get(url, headers=headers, timeout=15)
            
            if response.status_code == 200:
                return response.json()
            
            elif response.status_code == 429:
                # Serwer mówi: STOP. Pobieramy informację ile sekund czekać
                retry_after = response.headers.get("Retry-After")
                if retry_after and retry_after.isdigit():
                    sleep_time = int(retry_after)
                else:
                    # Wykładnicze opóźnienie (Exponential Backoff) + losowy "jitter" zapobiegający uderzeniu wszystkich wątków naraz
                    sleep_time = (base_delay ** attempt) + random.uniform(0.5, 1.5)
                
                time.sleep(sleep_time)
                continue  # Przejdź do kolejnej próby (spróbuj jeszcze raz później)
                
            else:
                # Inne błędy (np. 404 - brak produktu w bazie) -> nie ma sensu ponawiać
                break
        except (requests.exceptions.RequestException, Exception):
            # Błędy sieciowe/timeouty -> poczekaj chwilę i spróbuj ponownie
            time.sleep(3)
            
    return None

# --- LOGIKA PRZETWARZANIA JEDNEGO WIERSZA (URUCHAMIANA W WĄTKU) ---

def process_single_row(index, row, selected_cols, selected_idents, api_key, session):
    """
    Przetwarza jeden wiersz tabeli zgodnie z kaskadową logiką priorytetów.
    Funkcja jest thread-safe (bezpieczna dla wielowątkowości).
    """
    found_data = None
    matched_value = None
    checked_vals_in_row = set()
    
    # Kaskada po kolumnach
    for col in selected_cols:
        val = str(row[col]).split('.')[0].strip() if pd.notnull(row[col]) else ""
        if not val or val.lower() == 'nan':
            continue
        
        checked_vals_in_row.add(val)
        
        # Kaskada po identyfikatorach
        for ident in selected_idents:
            
            # --- OBSŁUGA KODU EPREL ---
            if ident == "Kod EPREL" and val.isdigit():
                # Bezpieczny odczyt/zapis z Cache pod blokadą systemową
                with cache_lock:
                    in_cache = val in cache_eprel
                    if in_cache:
                        found_data = cache_eprel[val]
                
                if not in_cache:
                    # Jeśli nie ma w cache, odpytujemy API (poza blokadą, żeby nie blokować innych wątków!)
                    url = f"https://eprel.ec.europa.eu/api/product/{val}"
                    headers = {"Authorization": f"Bearer {api_key}", "Accept": "application/json"}
                    res_data = call_eprel_api_with_retry(url, headers, session)
                    
                    with cache_lock:
                        cache_eprel[val] = res_data
                    found_data = res_data
                    time.sleep(0.1) # Lekki bezpiecznik dla stabilności bazy
            
            # --- OBSŁUGA KODU EAN ---
            elif ident == "EAN" and val.isdigit() and len(val) in [8, 12, 13, 14]:
                with cache_lock:
                    in_cache = val in cache_ean
                    if in_cache:
                        found_data = cache_ean[val]
                        
                if not in_cache:
                    url = f"https://eprel.ec.europa.eu/api/product/gtin/{val}"
                    headers = {"Authorization": f"Bearer {api_key}", "Accept": "application/json"}
                    res_data = call_eprel_api_with_retry(url, headers, session)
                    
                    with cache_lock:
                        cache_ean[val] = res_data
                    found_data = res_data
                    time.sleep(0.1)
            
            if found_data:
                matched_value = val
                break  # Sukces identyfikatora -> przerwij sprawdzanie kolejnych dla tej wartości
                
        if found_data:
            break  # Sukces kolumny -> przerwij sprawdzanie kolejnych kolumn dla tego wiersza
            
    # Zmapowanie pobranych danych strukturalnych
    if found_data:
        energy_class = found_data.get("energyClass", "N/A")
        model_identifier = found_data.get("modelIdentifier", "N/A")
        real_id = found_data.get("registrationNumber") or found_data.get("eprelRegistrationNumber") or ""
        product_group = found_data.get("productGroup", "N/A")
        pdf_link = f"https://eprel.ec.europa.eu/labels/{product_group}/Label_{real_id}.pdf" if (product_group != "N/A" and real_id) else "Brak grupy"
        
        return index, model_identifier, real_id, product_group, energy_class, pdf_link, True, checked_vals_in_row
    else:
        return index, "Brak danych", "Brak danych", "Brak danych", "Brak danych", "Brak danych", False, checked_vals_in_row

# --- UI STREAMLIT ---

st.title("⚡ EPRELap_KTA")
st.markdown("""
Aplikacja służy do odpytywania bazy EPREL o informacje o produkcie na podstawie identyfikatora, 
którym jest **Kod EPREL** lub **EAN**. 
_Wersja przemysłowa: zoptymalizowana pod kątem wielowątkowości (Concurrent Async) oraz odporności na blokady IP (Auto-Retry 429)._
""")

try:
    API_KEY = st.secrets["EPREL_API_KEY"]
except Exception:
    st.error("Błąd: Nie znaleziono klucza 'EPREL_API_KEY' w Secrets!")
    st.stop()

uploaded_file = st.file_uploader("Załaduj plik Excel (.xlsx)", type=["xlsx"])

if uploaded_file:
    df_in = pd.read_excel(uploaded_file)
    
    st.subheader("📋 Podgląd danych (Sample)")
    st.dataframe(df_in.head(5), use_container_width=True)
    
    st.divider()
    
    st.subheader("⚙️ Konfiguracja masowego odpytywania bazy")
    selected_cols = st.multiselect(
        "Wybierz kolumnę/kolumny z pliku bazowego (kolejność wyboru określa PRIORYTET):",
        options=df_in.columns
    )
    
    selected_idents = st.multiselect(
        "Wybierz typ identyfikatora / zapytania:",
        options=["Kod EPREL", "EAN"],
        default=["Kod EPREL"]
    )
    
    st.divider()
    
    if st.button("Uruchom szybkie przetwarzanie wielowątkowe"):
        if not selected_cols:
            st.error("Błąd: Musisz wybrać minimum jedną kolumnę!")
        elif not selected_idents:
            st.error("Błąd: Musisz wybrać minimum jeden typ identyfikatora!")
        else:
            # Czyszczenie globalnego cache przed nowym uruchomieniem
            with cache_lock:
                cache_eprel.clear()
                cache_ean.clear()
                
            total_rows = len(df_in)
            
            # Listy przygotowane na wyniki (z zachowaniem indeksów)
            results_map = {}
            unique_checked_values = set()
            unique_success_values = set()
            total_rows_success = 0
            
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            # Tworzymy współdzieloną sesję HTTP dla optymalizacji połączeń TCP
            http_session = requests.Session()
            
            # --- SEKCJA WIELOWĄTKOWOŚCI (THREAD POOL) ---
            # max_workers=3 oznacza, że maksymalnie 3 zapytania lecą do UE w tej samej milisekundzie
            with concurrent.futures.ThreadPoolExecutor(max_workers=3) as executor:
                # Wrzucamy wszystkie wiersze do puli zadań robotników
                futures = {
                    executor.submit(
                        process_single_row, idx, row, selected_cols, selected_idents, API_KEY, http_session
                    ): idx for idx, row in df_in.iterrows()
                }
                
                # Odbieramy wyniki dynamicznie, w miarę jak wątki kończą pracę
                for count, future in enumerate(concurrent.futures.as_completed(futures)):
                    try:
                        idx, model, r_id, group, r_class, link, success, checked_vals = future.result()
                        
                        # Zapisujemy wynik pod odpowiednim indeksem
                        results_map[idx] = {
                            "EPREL_Model": model,
                            "EPREL_ID": r_id,
                            "EPREL_Grupa": group,
                            "EPREL_Klasa": r_class,
                            "EPREL_Link_PDF": link
                        }
                        
                        # Statystyki do raportu
                        unique_checked_values.update(checked_vals)
                        if success:
                            total_rows_success += 1
                            # Pobieramy ostatnią udaną wartość z tego wiersza do statystyk unikalnych
                            if checked_vals:
                                unique_success_values.add(list(checked_vals)[-1])
                                
                    except Exception as e:
                        st.error(f"Wystąpił nieoczekiwany błąd krytyczny w wątku: {e}")
                    
                    # Aktualizacja paska postępu w UI (zoptymalizowana: co 1% lub przy końcu)
                    if count % max(1, total_rows // 100) == 0 or count == total_rows - 1:
                        progress_bar.progress((count + 1) / total_rows)
                        status_text.text(f"Ukończono wierszy: {count + 1} z {total_rows}")
            
            status_text.empty()
            
            # Przekształcenie mapy wyników na posortowaną listę odpowiadającą kolejności w Excelu
            ordered_results = [results_map[i] for i in range(total_rows)]
            df_results = pd.DataFrame(ordered_results)
            
            # Połączenie pliku bazowego z nowymi kolumnami
            df_out = pd.concat([df_in, df_results], axis=1)
            
            st.success("Wielowątkowe przetwarzanie ogromnej bazy zakończone sukcesem!")
            
            # --- RAPORT KOŃCOWY ---
            st.subheader("📊 Raport z przetwarzania")
            col_rep1, col_rep2, col_rep3, col_rep4 = st.columns(4)
            with col_rep1:
                st.metric("Przeanalizowane kody (unikalne)", len(unique_checked_values))
            with col_rep2:
                st.metric("Sukces (unikalne produkty)", len(unique_success_values))
            with col_rep3:
                st.metric("Uzupełnione wiersze w Excelu", total_rows_success)
            with col_rep4:
                api_calls = len(cache_eprel) + len(cache_ean)
                saved_calls = total_rows - api_calls
                st.metric("Faktyczne zapytania API", api_calls, delta=f"-{saved_calls} dzięki CACHE")
                
            # Prezentacja i pobieranie danych
            st.subheader("📥 Podgląd zaktualizowanego pliku")
            st.dataframe(
                df_out.head(100), # Pokazujemy pierwsze 100 wierszy dla płynności przeglądarki
                column_config={"EPREL_Link_PDF": st.column_config.LinkColumn("Link PDF")},
                use_container_width=True
            )
            
            buf_excel = io.BytesIO()
            with pd.ExcelWriter(buf_excel, engine='xlsxwriter') as writer:
                df_out.to_excel(writer, index=False)
            
            st.download_button(
                label="📥 Pobierz kompletny raport Excel (XLSX)",
                data=buf_excel.getvalue(),
                file_name="eprel_mega_raport_wielowatkowy.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
