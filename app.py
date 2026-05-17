import streamlit as st
import pandas as pd
from docx import Document
import io
import re

# Ustawienia strony Streamlit
st.set_page_config(page_title="Konwerter Scenariuszy UAT Comfortel", layout="wide")

st.title("📄 Profesjonalny Konwerter Scenariuszy UAT (SORT.B3S)")
st.write("Wersja silnika: **v3.0-Final**. Pełna obsługa nieregularnych układów tabel i scalonych komórek dokumentacji Comfortel.")

# Oficjalna struktura kolumn dla pliku Excel
TARGET_COLUMNS = [
    "Scenariusze testowe", "Moduł", "Pełny Nr wymagania", "Opis wymagania", 
    "Zakres wyłączeń", "Nr scenariusza", "Nazwa scenariusza", "Cel", 
    "Warunki wstępne", "LP", "Kroki testowe", "Oczekiwany rezultat", 
    "Wynik testu podczas odbioru", "Kategoria błędu", "Uwagi podczas odbioru"
]

def clean_text(text):
    """Dokładne czyszczenie tekstu z podwójnych spacji, znaków końca linii i tabulatorów."""
    if text is None:
        return ""
    return re.sub(r'\s+', ' ', str(text)).strip()

def parse_docx_final(uploaded_file):
    doc = Document(io.BytesIO(uploaded_file.read()))
    
    # 1. Próba automatycznego wykrycia Tytułu Dokumentu (np. [SORT.BSS | Obsługa Umów])
    main_title = ""
    for p in doc.paragraphs[:20]:
        t = p.text.strip()
        if "[SORT" in t or "Scenariusze testowe" in t:
            main_title += " " + t
    main_title = clean_text(main_title)
    if not main_title:
        main_title = "Scenariusze testowe systemu SORT.B3S"

    all_rows = []
    
    # Inicjalizacja kontekstu (bufora) dla pól nagłówkowych
    context = {col: "" for col in TARGET_COLUMNS}
    context["Scenariusze testowe"] = main_title

    # 2. Główna pętla przechodząca przez KAŻDĄ tabelę w dokumencie
    for table in doc.tables:
        
        # Przechodzimy wiersz po wierszu
        for row_idx, row in enumerate(table.rows):
            # Pobieramy oczyszczone teksty ze wszystkich komórek w tym wierszu
            cells_text = [clean_text(cell.text) for cell in row.cells]
            
            # Pomijamy puste wiersze
            if not any(cells_text):
                continue
                
            # --- SEKCJA A: WYKRYWANIE I AKTUALIZACJA KONTEKSTU (Pary Klucz-Wartość) ---
            # Sprawdzamy zawartość komórek pod kątem metadanych scenariusza/wymagania
            for idx, cell_txt in enumerate(cells_text):
                cell_lower = cell_txt.lower()
                
                # Wykrywanie Modułu
                if cell_lower == "moduł" and idx + 1 < len(cells_text):
                    context["Moduł"] = cells_text[idx + 1]
                # Wykrywanie Numeru Wymagania
                elif "nr wymagania" in cell_lower and idx + 1 < len(cells_text):
                    context["Pełny Nr wymagania"] = cells_text[idx + 1]
                # Wykrywanie Opisu Wymagania
                elif "opis wymagania" in cell_lower and idx + 1 < len(cells_text):
                    context["Opis wymagania"] = cells_text[idx + 1]
                # Wykrywanie Celu
                elif cell_lower == "cel" and idx + 1 < len(cells_text):
                    context["Cel"] = cells_text[idx + 1]
                # Wykrywanie Warunków wstępnych
                elif "warunki wstępne" in cell_lower and idx + 1 < len(cells_text):
                    context["Warunki wstępne"] = cells_text[idx + 1]
                
                # Wykrywanie Nagłówka Scenariusza (szukamy znaku # np. #TAKC_001)
                if cell_txt.startswith("#"):
                    match = re.match(r"^#\s*([A-Za-z0-9_]+)\s*[\–\-\—\:\.]\s*(.*)", cell_txt)
                    if match:
                        context["Nr scenariusza"] = match.group(1).strip()
                        context["Nazwa scenariusza"] = match.group(2).strip()
                    else:
                        context["Nr scenariusza"] = cell_txt.replace("#", "").strip()
                        context["Nazwa scenariusza"] = ""

            # --- SEKCJA B: WYKRYWANIE I PARSOWANIE WIERSZY Z KROKAMI ---
            # Sprawdzamy czy bieżący wiersz jest wierszem danych (zaczyna się od cyfry np. "1", "2")
            # ORAZ czy wiersz posiada odpowiednią liczbę kolumn (Comfortel stosuje zazwyczaj układ: LP, Dane, Opis, Rezultat)
            if len(cells_text) >= 3:
                first_cell = cells_text[0]
                
                # Sprawdzamy czy pierwsza komórka to liczba (Identyfikator kroku)
                if first_cell.isdigit():
                    lp_val = first_cell
                    
                    # W dokumentacji Comfortel:
                    # Jeśli mamy 3 kolumny: [LP, Opis kroku, Oczekiwany rezultat]
                    # Jeśli mamy 4 kolumny: [LP, Dane, Opis kroku, Oczekiwany rezultat]
                    if len(cells_text) == 3:
                        opis_val = cells_text[1]
                        rez_val = cells_text[2]
                    else:
                        # Wersja domyślna dla 4 kolumn (kolumna indeks 1 to zazwyczaj puste 'Dane')
                        opis_val = cells_text[2]
                        rez_val = cells_text[3]
                        
                    # Ignorujemy techniczne nagłówki powtórzone przez Worda
                    if "opis" in opis_val.lower() or "oczekiwany" in rez_val.lower():
                        continue
                        
                    # Zabezpieczenie: Dodajemy tylko jeśli krok zawiera jakąś treść
                    if opis_val or rez_val:
                        # Kopiujemy aktualny stan słownika i nadpisujemy wartościami kroku
                        flat_row = context.copy()
                        flat_row["LP"] = lp_val
                        flat_row["Kroki testowe"] = opis_val
                        flat_row["Oczekiwany rezultat"] = rez_val
                        
                        all_rows.append(flat_row)

    # 3. Konwersja zebranej listy słowników do formatu DataFrame
    if all_rows:
        df = pd.DataFrame(all_rows)
    else:
        df = pd.DataFrame(columns=TARGET_COLUMNS)
        
    # Upewniamy się, że w tabeli znajdą się puste kolumny wymagane w strukturze końcowej
    for col in TARGET_COLUMNS:
        if col not in df.columns:
            df[col] = ""
            
    # Zwracamy gotową tabelę w precyzyjnie zdefiniowanej kolejności kolumn
    return df[TARGET_COLUMNS]

# --- INTERFEJS UŻYTKOWNIKA (STREAMLIT) ---
uploaded_file = st.file_uploader("Wgraj plik scenariuszy Word (.docx)", type=["docx"])

if uploaded_file is not None:
    with st.spinner("Dekodowanie struktury tabel Comfortel..."):
        try:
            df_result = parse_docx_final(uploaded_file)
            
            if not df_result.empty:
                st.success(f"Sukces! Silnik v3.0 pomyślnie wyodrębnił i spłaszczył {len(df_result)} kroków testowych.")
                
                # Statystyki KPI na górze ekranu
                c1, c2, c3 = st.columns(3)
                with c1:
                    st.metric("Wygenerowane wiersze (kroki)", len(df_result))
                with c2:
                    st.metric("Wykryte scenariusze (#)", df_result["Nr scenariusza"].nunique())
                with c3:
                    st.metric("Zmapowane wymagania", df_result["Pełny Nr wymagania"].nunique())

                # Podgląd tabeli na żywo w Streamlit
                st.subheader("👀 Podgląd spłaszczonej tabeli przed pobraniem")
                st.dataframe(df_result, use_container_width=True)
                
                # Przygotowanie pobierania pliku Excel (.xlsx)
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    df_result.to_excel(writer, index=False, sheet_name='Scenariusze UAT')
                buffer.seek(0)
                
                st.download_button(
                    label="📥 Pobierz gotowy arkusz Excel (.xlsx)",
                    data=buffer,
                    file_name="spłaszczone_scenariusze_uat.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.error("Błąd: Silnik przeskanował cały dokument, ale nie znalazł wierszy rozpoczynających się od numerów kroków (cyfr) wewnątrz tabel. Upewnij się, że wgrałeś poprawny plik .docx.")
        
        except Exception as e:
            st.error(f"Wystąpił nieoczekiwany błąd przetwarzania: {e}")
