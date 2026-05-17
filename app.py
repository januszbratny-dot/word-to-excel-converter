import streamlit as st
import pandas as pd
from docx import Document
import io
import re

# Ustawienia strony Streamlit
st.set_page_config(page_title="Konwerter Scenariuszy UAT Comfortel", layout="wide")

st.title("📄 Zaawansowany Konwerter Scenariuszy UAT (Comfortel)")
st.write("Wersja silnika zoptymalizowana pod kątem nieregularnych i scalonych tabel dokumentacji SORT.")

# Definicja docelowych kolumn strukturalnych (zgodnie z Twoim wymaganiem)
TARGET_COLUMNS = [
    "Scenariusze testowe", "Moduł", "Pełny Nr wymagania", "Opis wymagania", 
    "Zakres wyłączeń", "Nr scenariusza", "Nazwa scenariusza", "Cel", 
    "Warunki wstępne", "LP", "Kroki testowe", "Oczekiwany rezultat", 
    "Wynik testu podczas odbioru", "Kategoria błędu", "Uwagi podczas odbioru"
]

def clean_text(text):
    """Pomocnicza funkcja czyszcząca tekst z białych znaków i zbędnych linii."""
    if text is None:
        return ""
    return re.sub(r'\s+', ' ', str(text)).strip()

def parse_docx_robust(uploaded_file):
    doc = Document(io.BytesIO(uploaded_file.read()))
    
    # Próba wyciągnięcia tytułu głównego z pierwszych akapitów dokumentu
    main_title = ""
    for p in doc.paragraphs[:15]:  # Sprawdzamy pierwsze 15 akapitów
        txt = p.text.strip()
        if "[SORT" in txt or "Scenariusze testowe" in txt:
            main_title += " " + txt
    
    main_title = clean_text(main_title)
    if not main_title:
        main_title = "Scenariusze testowe systemu SORT.B3S"

    all_rows = []
    
    # Inicjalizacja globalnego kontekstu (bufora) danych nagłówkowych
    context = {col: "" for col in TARGET_COLUMNS}
    context["Scenariusze testowe"] = main_title

    # Bezpieczne pobieranie wszystkich elementów w kolejności występowania
    # (Pętla przechodzi przez wszystkie akapity oraz tabele jako logiczny ciąg)
    for element in doc.element.body:
        # --- OBSŁUGA AKAPITÓW (Tekst bezpośredni) ---
        if element.tag.endswith('p'):
            p_text = ""
            # Mapowanie elementu XML na obiekt Paragraph python-docx
            p_objs = [p for p in doc.paragraphs if p._element == element]
            if p_objs:
                p_text = p_objs[0].text.strip()
            
            # Poszukiwanie linii scenariusza np. #TAKC_001 lub # TAKC_001
            if p_text.startswith("#"):
                # Wyrażenie regularne akceptuje różne myślniki, spacje i formaty
                match = re.match(r"^#\s*([A-Za-z0-9_]+)\s*[\–\-\—\:]\s*(.*)", p_text)
                if match:
                    context["Nr scenariusza"] = clean_text(match.group(1))
                    context["Nazwa scenariusza"] = clean_text(match.group(2))
                else:
                    context["Nr scenariusza"] = clean_text(p_text.replace("#", ""))
                    context["Nazwa scenariusza"] = ""

        # --- OBSŁUGA TABEL ---
        elif element.tag.endswith('tbl'):
            t_objs = [t for t in doc.tables if t._element == element]
            if not t_objs:
                continue
            table = t_objs[0]
            
            if len(table.rows) == 0:
                continue

            # Odczytanie pierwszego wiersza w celu identyfikacji typu tabeli
            first_row_cells = [clean_text(cell.text).lower() for cell in table.rows[0].cells]
            
            # Flagi identyfikacyjne
            is_steps_table = False
            
            # Sprawdzenie czy to tabela z krokami testowymi
            has_lp = any("lp" in c or "l.p" in c for c in first_row_cells)
            has_krok_or_opis = any("krok" in c or "opis" in c for c in first_row_cells)
            has_rezultat = any("rezultat" in c or "wynik" in c or "oczekiwany" in c for c in first_row_cells)
            
            if has_lp and (has_krok_or_opis or has_rezultat):
                is_steps_table = True

            if is_steps_table:
                # Dynamiczne mapowanie indeksów kolumn (odporność na kolumnę 'Dane')
                col_map = {"lp": None, "opis": None, "rezultat": None}
                for idx, cell_text in enumerate(first_row_cells):
                    if "lp" in cell_text or "l.p" in cell_text:
                        col_map["lp"] = idx
                    elif "opis" in cell_text or "krok" in cell_text:
                        col_map["opis"] = idx
                    elif "rezultat" in cell_text or "wynik" in cell_text or "oczekiwany" in cell_text:
                        col_map["rezultat"] = idx

                # Iteracja po wierszach z krokami (od drugiego wiersza)
                for row in table.rows[1:]:
                    # Zabezpieczenie przed niepełnymi lub pustymi wierszami ze scaleniami
                    cells_text = [clean_text(c.text) for c in row.cells]
                    
                    try:
                        lp_val = cells_text[col_map["lp"]] if col_map["lp"] is not None and col_map["lp"] < len(cells_text) else ""
                        opis_val = cells_text[col_map["opis"]] if col_map["opis"] is not None and col_map["opis"] < len(cells_text) else ""
                        rez_val = cells_text[col_map["rezultat"]] if col_map["rezultat"] is not None and col_map["rezultat"] < len(cells_text) else ""
                    except IndexError:
                        continue
                    
                    # Ignorujemy wiersze będące nagłówkami sekcji wewnątrz tabeli (np. ponowny napis KROKI TESTOWE)
                    if "kroki" in lp_val.lower() or "opis" in lp_val.lower():
                        continue

                    if lp_val or opis_val:
                        # Kopiujemy bieżący stan kontekstu i doklejamy dane kroku
                        flat_row = context.copy()
                        flat_row["LP"] = lp_val
                        flat_row["Kroki testowe"] = opis_val
                        flat_row["Oczekiwany rezultat"] = rez_val
                        
                        all_rows.append(flat_row)
            
            else:
                # TABELA NAGŁÓWKOWA / FORMULARZOWA (Klucz -> Wartość)
                # Parsujemy wiersze i uzupełniamy kontekst dla kolejnych tabel kroków
                for row in table.rows:
                    cells = [clean_text(c.text) for c in row.cells]
                    if len(cells) >= 2:
                        key = cells[0].lower()
                        val = cells[1]
                        
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
                        # Dodatkowy warunek, jeśli nazwa scenariusza pojawi się w tabeli formularza zamiast tekstu #
                        elif "nazwa scenariusza" in key:
                            context["Nazwa scenariusza"] = val
                        elif "nr scenariusza" in key:
                            context["Nr scenariusza"] = val

    # Tworzenie wynikowego obiektu DataFrame
    if all_rows:
        df = pd.DataFrame(all_rows)
    else:
        df = pd.DataFrame(columns=TARGET_COLUMNS)
    
    # Gwarancja istnienia i poprawnej kolejności wszystkich wymaganych kolumn
    for col in TARGET_COLUMNS:
        if col not in df.columns:
            df[col] = ""
            
    return df[TARGET_COLUMNS]

# --- INTERFEJS UŻYTKOWNIKA STREAMLIT ---
uploaded_file = st.file_uploader("Wgraj plik scenariuszy Word (.docx)", type=["docx"])

if uploaded_file is not None:
    with st.spinner("Przetwarzanie dokumentu przy użyciu elastycznego silnika..."):
        try:
            df_result = parse_docx_robust(uploaded_file)
            
            if not df_result.empty:
                st.success(f"Sukces! Wykryto i spłaszczono {len(df_result)} kroków testowych.")
                
                # Karty statystyk (KPI)
                c1, c2, c3 = st.columns(3)
                with c1:
                    st.metric("Przetworzone kroki", len(df_result))
                with c2:
                    st.metric("Unikalne scenariusze", df_result["Nr scenariusza"].nunique())
                with c3:
                    st.metric("Powiązane wymagania", df_result["Pełny Nr wymagania"].nunique())

                # Sekcja podglądu tabeli na żywo
                st.subheader("👀 Podgląd wygenerowanej płaskiej struktury tabeli")
                st.dataframe(df_result, use_container_width=True)
                
                # Tworzenie strumienia binarnego do pobrania pliku Excel (.xlsx)
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    df_result.to_excel(writer, index=False, sheet_name='Scenariusze UAT')
                buffer.seek(0)
                
                st.download_button(
                    label="📥 Pobierz wygenerowany arkusz Excel (.xlsx)",
                    data=buffer,
                    file_name="scenariusze_flat_table.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.warning("Silnik zakończył pracę, lecz wynikowa tabela jest pusta. Sprawdź, czy plik zawiera tabele spełniające kryteria kolumn (LP + Opis/Rezultat).")
        
        except Exception as e:
            st.error(f"Błąd krytyczny podczas przetwarzania struktury dokumentu: {e}")
