import streamlit as st
import pandas as pd
from docx import Document
import io

# Ustawienia strony Streamlit
st.set_page_config(page_title="Konwerter Word -> Excel", layout="wide")

st.title("📄 Zaawansowany Konwerter Scenariuszy Testowych")
st.write("Skrypt analizuje tabele w Wordzie jako pary klucz-wartość i płaszczy dane do docelowego formatu.")

# Docelowa struktura kolumn w Excelu
TARGET_COLUMNS = [
    "Scenariusze testowe", "Moduł", "Pełny Nr wymagania", "Opis wymagania", 
    "Zakres wyłączeń", "Nr scenariusza", "Nazwa scenariusza", "Cel", 
    "Warunki wstępne", "LP", "Kroki testowe", "Oczekiwany rezultat", 
    "Wynik testu podczas odbioru", "Kategoria błędu", "Uwagi podczas odbioru"
]

# Słownik mapowania: Jakie frazy z tabeli w Wordzie odpowiadają kolumnom w Excelu?
# WPISZ TUTAJ DOKŁADNE NAZWY / FRAGMENTY NAZW Z TWOJEGO PLIKU WORD
MAPPING_DICTIONARY = {
    "Scenariusze testowe": ["scenariusz testowy", "scenariusze", "projekt"],
    "Moduł": ["moduł", "modul", "obszar"],
    "Pełny Nr wymagania": ["pełny nr wymagania", "nr wymagania", "id wymagania", "wymaganie"],
    "Opis wymagania": ["opis wymagania", "wymaganie opis"],
    "Zakres wyłączeń": ["zakres wyłączeń", "wyłączenia", "wylaczenia"],
    "Nr scenariusza": ["nr scenariusza", "id scenariusza", "numer scenariusza"],
    "Nazwa scenariusza": ["nazwa scenariusza", "tytuł scenariusza"],
    "Cel": ["cel", "cel testu"],
    "Warunki wstępne": ["warunki wstępne", "warunki wstepne", "preklimatyzacja"],
    "LP": ["lp", "l.p.", "lp."],
    "Kroki testowe": ["kroki testowe", "krok", "opis kroku", "działanie"],
    "Oczekiwany rezultat": ["oczekiwany rezultat", "rezultat", "oczekiwany wynik"],
    "Wynik testu podczas odbioru": ["wynik testu", "wynik testu podczas odbioru", "status"],
    "Kategoria błędu": ["kategoria błędu", "kategoria", "błąd"],
    "Uwagi podczas odbioru": ["uwagi", "uwagi podczas odbioru", "komentarz"]
}

def find_column_name(cell_text):
    """Funkcja dopasowuje tekst z komórki Worda do oficjalnej kolumny na podstawie słownika mapowania."""
    clean_text = cell_text.strip().lower().replace(":", "") # Usunięcie dwukropków i wielkich liter
    for target_col, aliases in MAPPING_DICTIONARY.items():
        for alias in aliases:
            if alias in clean_text:
                return target_col
    return None

def parse_docx_to_dataframe(uploaded_file):
    doc = Document(io.BytesIO(uploaded_file.read()))
    all_rows_data = []
    
    # Słownik przechowujący stan globalny (pamięć podręczna dla pól nagłówkowych)
    # Dzięki temu "Moduł" czy "Cel" wpisany raz na górze tabeli, przypisze się do każdego kroku.
    current_context = {col: "" for col in TARGET_COLUMNS}

    for table_idx, table in enumerate(doc.tables):
        # Sprawdzamy orientację tabeli na podstawie pierwszego wiersza
        first_row = [cell.text.strip() for cell in table.rows[0].cells]
        
        # Słownik mapowania indeksów kolumn dla układu poziomej tabeli
        horizontal_mapping = {}
        is_horizontal_table = False
        
        # Sprawdź czy pierwszy wiersz zawiera nagłówki (np. LP, Kroki testowe)
        for idx, cell_text in enumerate(first_row):
            matched_col = find_column_name(cell_text)
            if matched_col:
                horizontal_mapping[idx] = matched_col
                # Jeśli znaleźliśmy typowo "tabelaryczne" nagłówki, traktujemy tabelę jako poziomą
                if matched_col in ["LP", "Kroki testowe", "Oczekiwany rezultat"]:
                    is_horizontal_table = True

        if is_horizontal_table:
            # --- UKŁAD POZIOMY (Tabela z krokami) ---
            for row in table.rows[1:]: # Pomijamy nagłówek
                row_cells = [cell.text.strip() for cell in row.cells]
                
                # Tworzymy nowy wiersz oparty na kontekście (dziedziczymy np. Nazwę scenariusza)
                row_data = current_context.copy()
                
                # Nadpisujemy wartościami z bieżącego wiersza tabeli kroków
                for idx, cell_value in enumerate(row_cells):
                    if idx in horizontal_mapping:
                        target_col = horizontal_mapping[idx]
                        row_data[target_col] = cell_value
                
                # Dodajemy wiersz tylko, jeśli zawiera przynajmniej krok lub LP (żeby unikać pustych linii)
                if row_data["Kroki testowe"] or row_data["LP"]:
                    all_rows_data.append(row_data)
        else:
            # --- UKŁAD PIONOWY LUB FORMULARZOWY ---
            # Szukamy par klucz-wartość wewnątrz tabeli (np. Kolumna 0 to Klucz, Kolumna 1 to Wartość)
            for row in table.rows:
                if len(row.cells) >= 2:
                    key_text = row.cells[0].text
                    value_text = row.cells[1].text
                    
                    matched_col = find_column_name(key_text)
                    if matched_col:
                        # Aktualizujemy kontekst (zapamiętujemy np. aktualny Cel/Moduł)
                        current_context[matched_col] = value_text.strip()

    # Zamiana listy słowników na czysty DataFrame pandas
    df = pd.DataFrame(all_rows_data)
    
    # W razie gdyby braki sprawiły, że nie ma jakichś kolumn, upewniamy się, że struktura jest kompletna
    for col in TARGET_COLUMNS:
        if col not in df.columns:
            df[col] = ""
            
    return df[TARGET_COLUMNS]

# --- INTERFEJS STREAMLIT ---
uploaded_file = st.file_uploader("Wybierz plik Word (.docx)", type=["docx"])

if uploaded_file is not None:
    with st.spinner("Trwa analizowanie struktury tabel..."):
        try:
            df_result = parse_docx_to_dataframe(uploaded_file)
            
            if not df_result.empty:
                st.success("Tabele zostały pomyślnie spłaszczone do postaci par klucz-wartość!")
                
                # Wyświetlenie statystyk
                st.metric(label="Liczba wygenerowanych wierszy (kroków)", value=len(df_result))
                
                # Podgląd tabeli
                st.subheader("👀 Podgląd gotowego arkusza")
                st.dataframe(df_result, use_container_width=True)
                
                # Przygotowanie pliku Excel do pobrania
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    df_result.to_excel(writer, index=False, sheet_name='Scenariusze_Spłaszczone')
                buffer.seek(0)
                
                st.download_button(
                    label="📥 Pobierz plik Excel (.xlsx)",
                    data=buffer,
                    file_name="scenariusze_wyjsciowe.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.warning("Nie udało się dopasować kluczy z Worda do wymaganych kolumn. Upewnij się, że nazwy pól w Wordzie są podobne do docelowych kolumn.")
                
        except Exception as e:
            st.error(f"Wystąpił błąd podczas mapowania danych: {e}")
