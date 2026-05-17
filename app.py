import streamlit as st
import pandas as pd
from docx import Document
import io

# Ustawienia strony Streamlit
st.set_page_config(page_title="Konwerter Word -> Excel", layout="wide")

st.title("📄 Konwerter Scenariuszy Testowych z Word do Excel")
st.write("Wgraj plik .docx zawierający tabele ze scenariuszami, sprawdź podgląd i pobierz gotowy arkusz Excel.")

# Zdefiniowanie docelowych kolumn
TARGET_COLUMNS = [
    "Scenariusze testowe", "Moduł", "Pełny Nr wymagania", "Opis wymagania", 
    "Zakres wyłączeń", "Nr scenariusza", "Nazwa scenariusza", "Cel", 
    "Warunki wstępne", "LP", "Kroki testowe", "Oczekiwany rezultat", 
    "Wynik testu podczas odbioru", "Kategoria błędu", "Uwagi podczas odbioru"
]

def parse_docx_to_dataframe(uploaded_file):
    # Wczytanie dokumentu z pamięci binarnej
    doc = Document(io.BytesIO(uploaded_file.read()))
    
    all_data = []
    
    # Iteracja po wszystkich tabelach w dokumencie
    for table in doc.tables:
        # Pomijamy nagłówek tabeli (zakładamy, że wiersz 0 to nagłówki)
        for row in table.rows[1:]:
            # Pobieramy tekst z każdej komórki w wierszu
            row_text = [cell.text.strip() for cell in row.cells]
            
            # --- TUTAJ DOPASUJ LOGIKĘ MAPOWANIA ---
            # Jeśli Twoja tabela w Wordzie ma mniej kolumn niż docelowy Excel,
            # musisz ręcznie przypisać wartości do odpowiednich miejsc.
            # Przykład poniżej zakłada, że wiersz ma dokładnie tyle samo kolumn:
            
            if len(row_text) >= len(TARGET_COLUMNS):
                # Obcinamy do długości docelowej, jeśli jest za długa
                all_data.append(row_text[:len(TARGET_COLUMNS)])
            else:
                # Dopełniamy pustymi wartościami, jeśli wiersz jest za krótki
                padded_row = row_text + [""] * (len(TARGET_COLUMNS) - len(row_text))
                all_data.append(padded_row)
                
    # Tworzenie DataFrame z dopasowanymi kolumnami
    df = pd.DataFrame(all_data, columns=TARGET_COLUMNS)
    return df

# Komponent do wgrywania pliku
uploaded_file = st.file_uploader("Wybierz plik Word (.docx)", type=["docx"])

if uploaded_file is not None:
    with st.spinner("Przetwarzanie pliku..."):
        try:
            # Konwersja tabeli
            df_result = parse_docx_to_dataframe(uploaded_file)
            
            if not df_result.empty:
                st.success("Plik przetworzony pomyślnie!")
                
                # --- PODGLĄD TABELI ---
                st.subheader("👀 Podgląd generowanej tabeli (płaska struktura)")
                st.dataframe(df_result, use_container_width=True)
                
                # --- GENEROWANIE EXCELA DO POBRANIA ---
                # Zapisujemy Excel do pamięci (buffer), aby Streamlit mógł go pobrać
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    df_result.to_excel(writer, index=False, sheet_name='Scenariusze')
                
                buffer.seek(0)
                
                # Przycisk pobierania
                st.download_button(
                    label="📥 Pobierz plik Excel (.xlsx)",
                    data=buffer,
                    file_name="scenariusze_testowe.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.warning("Nie znaleziono żadnych tabel w przesłanym dokumencie.")
                
        except Exception as e:
            st.error(f"Wystąpił błąd podczas przetwarzania pliku: {e}")
