import streamlit as st
import pandas as pd
from docx import Document
import io
import re

# Ustawienia strony Streamlit
st.set_page_config(page_title="Konwerter Scenariuszy UAT Comfortel", layout="wide")

st.title("📄 Zaawansowany Konwerter Scenariuszy UAT (Comfortel)")
st.write("Wersja silnika: **v4.0-Fix**. Poprawiono wykrywanie pierwszego kroku oraz dodano kolumnę 'Dane'.")

# Zaktualizowana struktura kolumn w Excelu (dodano "Dane" pomiędzy LP a Kroki testowe)
TARGET_COLUMNS = [
    "Scenariusze testowe", "Moduł", "Pełny Nr wymagania", "Opis wymagania", 
    "Zakres wyłączeń", "Nr scenariusza", "Nazwa scenariusza", "Cel", 
    "Warunki wstępne", "LP", "Dane", "Kroki testowe", "Oczekiwany rezultat", 
    "Wynik testu podczas odbioru", "Kategoria błędu", "Uwagi podczas odbioru"
]

def clean_text(text):
    """Dokładne czyszczenie tekstu z podwójnych spacji i znaków końca linii."""
    if text is None:
        return ""
    return re.sub(r'\s+', ' ', str(text)).strip()

def parse_docx_v4(uploaded_file):
    doc = Document(io.BytesIO(uploaded_file.read()))
    
    # Próba automatycznego wykrycia Tytułu Dokumentu
    main_title = ""
    for p in doc.paragraphs[:20]:
        t = p.text.strip()
        if "[SORT" in t or "Scenariusze testowe" in t:
            main_title += " " + t
    main_title = clean_text(main_title)
    if not main_title:
        main_title = "Scenariusze testowe systemu SORT.B3S"

    all_rows = []
    
    # Inicjalizacja kontekstu danych nagłówkowych
    context = {col: "" for col in TARGET_COLUMNS}
    context["Scenariusze testowe"] = main_title

    # Iteracja po wszystkich tabelach dokumentu
    for table in doc.tables:
        col_map = {"lp": None, "dane": None, "opis": None, "rezultat": None}
        is_steps_table = False
        
        # 1. NAJPZÓD SZUKAMY NAGŁÓWKA TABELI KROKÓW, ABY POZNAĆ INDEKSY KOLUMN
        for row_idx, row in enumerate(table.rows):
            cells_text = [clean_text(cell.text).lower() for cell in row.cells]
            
            # Flaga rozpoznania nagłówka
            has_lp = any(c == "lp" or c == "l.p." or c == "l.p" for c in cells_text)
            has_opis = any("opis" in c or "krok" in c for c in cells_text)
            
            if has_lp and has_opis:
                is_steps_table = True
                # Mapujemy dokładne pozycje indeksów kolumn w tym konkretnym obiekcie tabeli
                for idx, text in enumerate(cells_text):
                    if text in ["lp", "l.p.", "l.p"]:
                        col_map["lp"] = idx
                    elif "dane" in text:
                        col_map["dane"] = idx
                    elif "opis" in text or "krok" in text:
                        col_map["opis"] = idx
                    elif "rezultat" in text or "wynik" in text or "oczekiwany" in text:
                        col_map["rezultat"] = idx
                break # Znaleźliśmy nagłówek, przerywamy wyszukiwanie struktury struktury

        # 2. JEŚLI TO TABELA KROKÓW - PARSUJEMY DANE OD POCZĄTKU DO KOŃCA
        if is_steps_table:
            for row in table.rows:
                cells_raw = [clean_text(cell.text) for cell in row.cells]
                
                # Zabezpieczenie przed pustymi wierszami
                if not any(cells_raw):
                    continue
                
                # Sprawdzamy zawartość kolumny LP za pomocą indeksu z mapy nagłówka
                lp_idx = col_map["lp"]
                if lp_idx is not None and lp_idx < len(cells_raw):
                    lp_value = cells_raw[lp_idx]
                    
                    # Jeśli wiersz zawiera cyfrę w kolumnie LP (np. "1", "2", "10") - to jest to nasz krok!
                    if re.match(r"^\d+$", lp_value):
                        
                        # Dynamiczne wyciąganie wartości na podstawie mapy nagłówków
                        dane_idx = col_map["dane"]
                        opis_idx = col_map["opis"]
                        rez_idx = col_map["rezultat"]
                        
                        dane_val = cells_raw[dane_idx] if dane_idx is not None and dane_idx < len(cells_raw) else ""
                        opis_val = cells_raw[opis_idx] if opis_idx is not None and opis_idx < len(cells_raw) else ""
                        rez_val = cells_raw[rez_idx] if rez_idx is not None and rez_idx < len(cells_raw) else ""
                        
                        # Budujemy płaski wiersz i przypisujemy do TARGET_COLUMNS
                        flat_row = context.copy()
                        flat_row["LP"] = lp_value
                        flat_row["Dane"] = dane_val
                        flat_row["Kroki testowe"] = opis_val
                        flat_row["Oczekiwany rezultat"] = rez_val
                        
                        all_rows.append(flat_row)
                        
        else:
            # 3. JEŚLI TO NIE TABELA KROKÓW - TO TABELA METADANYCH (Klucz -> Wartość)
            for row in table.rows:
                cells = [clean_text(cell.text) for cell in row.cells]
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
                    elif "nazwa scenariusza" in key:
                        context["Nazwa scenariusza"] = val
                    elif "nr scenariusza" in key:
                        context["Nr scenariusza"] = val

                # Dodatkowe sprawdzenie, czy wewnątrz zwykłej tabeli nie ma hasha scenariusza (#TAKC_...)
                for cell_txt in cells:
                    if cell_txt.startswith("#"):
                        match = re.match(r"^#\s*([A-Za-z0-9_]+)\s*[\–\-\—\:\.]\s*(.*)", cell_txt)
                        if match:
                            context["Nr scenariusza"] = match.group(1).strip()
                            context["Nazwa scenariusza"] = match.group(2).strip()
                        else:
                            context["Nr scenariusza"] = cell_txt.replace("#", "").strip()
                            context["Nazwa scenariusza"] = ""

    # Tworzenie DataFrame
    if all_rows:
        df = pd.DataFrame(all_rows)
    else:
        df = pd.DataFrame(columns=TARGET_COLUMNS)
        
    # Zapewnienie kompletności kolumn
    for col in TARGET_COLUMNS:
        if col not in df.columns:
            df[col] = ""
            
    return df[TARGET_COLUMNS]

# --- SEKCJA INTERFEJSU STREAMLIT ---
uploaded_file = st.file_uploader("Wgraj plik scenariuszy Word (.docx)", type=["docx"])

if uploaded_file is not None:
    with st.spinner("Przetwarzanie danych przez zaktualizowany silnik v4.0..."):
        try:
            df_result = parse_docx_v4(uploaded_file)
            
            if not df_result.empty:
                st.success(f"Sukces! Poprawnie zmapowano {len(df_result)} kroków testowych z uwzględnieniem kolumny 'Dane'.")
                
                # Karty z metrykami KPI
                c1, c2, c3 = st.columns(3)
                with c1:
                    st.metric("Wyciągnięte kroki (Ogółem)", len(df_result))
                with c2:
                    st.metric("Liczba Scenariuszy (#)", df_result["Nr scenariusza"].nunique())
                with c3:
                    st.metric("Liczba Wymagań", df_result["Pełny Nr wymagania"].nunique())

                # Wyświetlenie podglądu tabeli
                st.subheader("👀 Podgląd zaktualizowanego arkusza (Zweryfikuj kolumnę 'Dane' i Krok 1)")
                st.dataframe(df_result, use_container_width=True)
                
                # Przygotowanie eksportu XLSX
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    df_result.to_excel(writer, index=False, sheet_name='Scenariusze_UAT_v4')
                buffer.seek(0)
                
                st.download_button(
                    label="📥 Pobierz poprawiony plik Excel (.xlsx)",
                    data=buffer,
                    file_name="scenariusze_uat_poprawione.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.error("Tabela wynikowa jest pusta. Upewnij się, że plik zawiera poprawnie sformatowane tabele kroków.")
        except Exception as e:
            st.error(f"Wystąpił błąd podczas parsowania: {e}")
