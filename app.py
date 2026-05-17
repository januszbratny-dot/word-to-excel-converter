import streamlit as st
import pandas as pd
from docx import Document
import io
import re

# Ustawienia strony Streamlit
st.set_page_config(page_title="Konwerter Scenariuszy UAT Comfortel", layout="wide")

st.title("📄 Profesjonalny Konwerter Scenariuszy UAT (Comfortel)")
st.write("Wersja silnika: **v5.0-Stable**. Pełna retencja danych nagłówkowych (brak pustych pól w spłaszczonej strukturze).")

# Docelowa struktura tabeli w Excelu
TARGET_COLUMNS = [
    "Scenariusze testowe", "Moduł", "Pełny Nr wymagania", "Opis wymagania", 
    "Zakres wyłączeń", "Nr scenariusza", "Nazwa scenariusza", "Cel", 
    "Warunki wstępne", "LP", "Dane", "Kroki testowe", "Oczekiwany rezultat", 
    "Wynik testu podczas odbioru", "Kategoria błędu", "Uwagi podczas odbioru"
]

def clean_text(text):
    """Czyszczenie tekstu ze zbędnych spacji, tabulatorów i znaków nowej linii."""
    if text is None:
        return ""
    return re.sub(r'\s+', ' ', str(text)).strip()

def parse_docx_v5(uploaded_file):
    doc = Document(io.BytesIO(uploaded_file.read()))
    
    # Próba automatycznego wykrycia Tytułu Dokumentu głównego
    main_title = ""
    for p in doc.paragraphs[:20]:
        t = p.text.strip()
        if "[SORT" in t or "Scenariusze testowe" in t:
            main_title += " " + t
    main_title = clean_text(main_title)
    if not main_title:
        main_title = "Scenariusze testowe systemu SORT.B3S"

    all_rows = []
    
    # Inicjalizacja trwałego kontekstu - wartości te będą pamiętane dopóki nie zmienią się na nowe
    context = {col: "" for col in TARGET_COLUMNS}
    context["Scenariusze testowe"] = main_title

    # Przechodzimy przez każdą tabelę w dokumencie Word
    for table in doc.tables:
        col_map = {"lp": None, "dane": None, "opis": None, "rezultat": None}
        is_steps_table = False
        
        # KROK 1: Skanowanie wierszy tabeli w poszukiwaniu nagłówka kroków testowych
        for row in table.rows:
            cells_text = [clean_text(cell.text).lower() for cell in row.cells]
            has_lp = any(c == "lp" or c == "l.p." or c == "l.p" for c in cells_text)
            has_opis = any("opis" in c or "krok" in c for c in cells_text)
            
            if has_lp and has_opis:
                is_steps_table = True
                # Mapowanie precyzyjnych indeksów kolumn
                for idx, text in enumerate(cells_text):
                    if text in ["lp", "l.p.", "l.p"]:
                        col_map["lp"] = idx
                    elif "dane" in text:
                        col_map["dane"] = idx
                    elif "opis" in text or "krok" in text:
                        col_map["opis"] = idx
                    elif "rezultat" in text or "wynik" in text or "oczekiwany" in text:
                        col_map["rezultat"] = idx
                break

        # KROK 2: Jeśli wykryto tabelę kroków - wyciągamy dane i łączymy z trwałym kontekstem
        if is_steps_table:
            for row in table.rows:
                cells_raw = [clean_text(cell.text) for cell in row.cells]
                
                if not any(cells_raw):
                    continue
                
                lp_idx = col_map["lp"]
                if lp_idx is not None and lp_idx < len(cells_raw):
                    lp_value = cells_raw[lp_idx]
                    
                    # Wyszukiwanie liczby (identyfikatora kroku) np. "1", "2"
                    if re.match(r"^\d+$", lp_value):
                        dane_idx = col_map["dane"]
                        opis_idx = col_map["opis"]
                        rez_idx = col_map["rezultat"]
                        
                        dane_val = cells_raw[dane_idx] if dane_idx is not None and dane_idx < len(cells_raw) else ""
                        opis_val = cells_raw[opis_idx] if opis_idx is not None and opis_idx < len(cells_raw) else ""
                        rez_val = cells_raw[rez_idx] if rez_idx is not None and rez_idx < len(cells_raw) else ""
                        
                        # Ignorujemy powtórzone nagłówki techniczne
                        if "opis" in opis_val.lower() or "oczekiwany" in rez_val.lower():
                            continue
                        
                        # Budujemy płaski wiersz na bazie TRWAŁEGO kontekstu
                        flat_row = context.copy()
                        flat_row["LP"] = lp_value
                        flat_row["Dane"] = dane_val
                        flat_row["Kroki testowe"] = opis_val
                        flat_row["Oczekiwany rezultat"] = rez_val
                        
                        all_rows.append(flat_row)
                        
        else:
            # KROK 3: Jeśli to NIE tabela kroków - to tabela metadanych. Aktualizujemy bufor kontekstu.
            for row in table.rows:
                cells = [clean_text(cell.text) for cell in row.cells]
                if len(cells) >= 2:
                    key = cells[0].lower()
                    val = cells[1]
                    
                    # Filtrujemy, aby upewnić się, że wartość nie jest identyczna jak klucz (błąd scalenia w Wordzie)
                    if key != val.lower():
                        if "moduł" in key or "modul" in key:
                            context["Moduł"] = val
                        elif "nr wymagania" in key or "numer wymagania" in key or "id wymagania" in key:
                            context["Pełny Nr wymagania"] = val
                        elif "opis wymagania" in key:
                            context["Opis wymagania"] = val
                        elif "cel" in key:
                            context["Cel"] = val
                        elif "warunki" in key:
                            context["Warunki wstępne"] = val
                        elif "nazwa scenariusza" in key:
                            context["Nazwa scenariusza"] = val
                        elif "nr scenariusza" in key:
                            context["Nr scenariusza"] = val

                # Przeszukanie komórek pod kątem nagłówka scenariusza typu: #TAKC_001
                for cell_txt in cells:
                    if cell_txt.startswith("#"):
                        match = re.match(r"^#\s*([A-Za-z0-9_]+)\s*[\–\-\—\:\.]\s*(.*)", cell_txt)
                        if match:
                            context["Nr scenariusza"] = match.group(1).strip()
                            context["Nazwa scenariusza"] = match.group(2).strip()
                        else:
                            context["Nr scenariusza"] = cell_txt.replace("#", "").strip()
                            context["Nazwa scenariusza"] = ""

    # Budowanie DataFrame z zebranych danych
    df = pd.DataFrame(all_rows) if all_rows else pd.DataFrame(columns=TARGET_COLUMNS)
    
    # Gwarancja istnienia wszystkich kolumn docelowych
    for col in TARGET_COLUMNS:
        if col not in df.columns:
            df[col] = ""
            
    return df[TARGET_COLUMNS]

# --- INTERFEJS UŻYTKOWNIKA STREAMLIT ---
uploaded_file = st.file_uploader("Wgraj plik scenariuszy Word (.docx)", type=["docx"])

if uploaded_file is not None:
    with st.spinner("Stabilne przetwarzanie struktury dokumentu v5.0..."):
        try:
            df_result = parse_docx_v5(uploaded_file)
            
            if not df_result.empty:
                st.success(f"Sukces! Poprawnie wygenerowano płaską strukturę tabeli. Liczba wierszy: {len(df_result)}.")
                
                # Wskaźniki jakości danych
                c1, c2, c3 = st.columns(3)
                with c1:
                    st.metric("Liczba kroków testowych", len(df_result))
                with c2:
                    st.metric("Zmapowane Scenariusze", df_result["Nr scenariusza"].nunique())
                with c3:
                    st.metric("Wypełnienie pól nadrzędnych", f"{int((df_result['Cel'] != '').mean() * 100)}%")

                # Tabela z podglądem na żywo
                st.subheader("👀 Podgląd kompletnej, spłaszczonej tabeli")
                st.dataframe(df_result, use_container_width=True)
                
                # Przygotowanie pobierania Excel
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    df_result.to_excel(writer, index=False, sheet_name='Scenariusze UAT')
                buffer.seek(0)
                
                st.download_button(
                    label="📥 Pobierz kompletny plik Excel (.xlsx)",
                    data=buffer,
                    file_name="scenariusze_uat_kompletne.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.error("Błąd: Wynikowa tabela jest pusta. Sprawdź poprawność pliku wejściowego.")
        except Exception as e:
            st.error(f"Wystąpił błąd krytyczny: {e}")
