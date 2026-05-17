import streamlit as st
import pandas as pd
from docx import Document
import io
import re

st.set_page_config(page_title="Konwerter Scenariuszy UAT", layout="wide")

st.title("📄 Dedykowany Konwerter Scenariuszy UAT (SORT.B3S)")
st.write("Aplikacja dostosowana do struktury dokumentacji Comfortel dla Etapu 3.")

# Definicja docelowych kolumn strukturalnych
TARGET_COLUMNS = [
    "Scenariusze testowe", "Moduł", "Pełny Nr wymagania", "Opis wymagania", 
    "Zakres wyłączeń", "Nr scenariusza", "Nazwa scenariusza", "Cel", 
    "Warunki wstępne", "LP", "Kroki testowe", "Oczekiwany rezultat", 
    "Wynik testu podczas odbioru", "Kategoria błędu", "Uwagi podczas odbioru"
]

def parse_strict_uat_docx(uploaded_file):
    doc = Document(io.BytesIO(uploaded_file.read()))
    
    # Wyciąganie tytułu głównego (z nagłówka dokumentu/pierwszych linii)
    main_title = ""
    for p in doc.paragraphs:
        if "Scenariusze testowe" in p.text or "[SORT" in p.text:
            main_title += " " + p.text.strip()
    main_title = main_title.strip() if main_title else "Scenariusze testowe SORT"

    all_rows = []
    
    # Globalny kontekst, który aktualizuje się podczas czytania dokumentu
    context = {col: "" for col in TARGET_COLUMNS}
    context["Scenariusze testowe"] = main_title

    # Analizujemy wszystkie elementy dokumentu w kolejności ich występowania
    for element in doc.element.body:
        # 1. Przetwarzanie PARAGRAFÓW (szukamy nagłówków scenariuszy typu #TAKC_001)
        if element.tag.endswith('p'):
            text = element.text if hasattr(element, 'text') else ""
            if not text:
                # Czasami trzeba pobrać tekst przez obiekt paragraph z docx
                p_obj = [p for p in doc.paragraphs if p._element == element]
                text = p_obj[0].text.strip() if p_obj else ""
            
            if text.startswith("#"):
                # Przykładowy format: "#TAKC_001 – Zmiana terminu montażu"
                match = re.match(r"^#([A-Za-z0-9_]+)\s*[\–\-]\s*(.*)", text)
                if match:
                    context["Nr scenariusza"] = match.group(1).strip()
                    context["Nazwa scenariusza"] = match.group(2).strip()
                else:
                    context["Nr scenariusza"] = text
                    context["Nazwa scenariusza"] = ""

        # 2. Przetwarzanie TABEL
        elif element.tag.endswith('tbl'):
            table_obj = [t for t in doc.tables if t._element == element]
            if not table_obj:
                continue
            table = table_obj[0]
            
            # Sprawdźmy, czy to tabela kroków (czy zawiera kolumny LP i Opis kroku/Oczekiwany rezultat)
            first_row_text = [cell.text.strip().lower() for cell in table.rows[0].cells]
            
            is_steps_table = any("lp" in cell or "l.p." in cell for cell in first_row_text) and \
                             any("krok" in cell or "opis" in cell for cell in first_row_text)

            if is_steps_table:
                # Mapowanie indeksów kolumn w tabeli kroków
                col_indices = {"lp": None, "opis": None, "rezultat": None}
                for idx, cell_text in enumerate(first_row_text):
                    if "lp" in cell_text or "l.p." in cell_text:
                        col_indices["lp"] = idx
                    elif "opis" in cell_text or "krok" in cell_text:
                        col_indices["opis"] = idx
                    elif "rezultat" in cell_text or "wynik" in cell_text or "oczekiwany" in cell_text:
                        col_indices["rezultat"] = idx
                
                # Iteracja po wierszach kroków testowych
                for row in table.rows[1:]:
                    cells = [c.text.strip() for c in row.cells]
                    
                    # Sprawdzenie zabezpieczające przed pustymi wierszami lub błędami indeksowania
                    try:
                        lp_val = cells[col_indices["lp"]] if col_indices["lp"] is not None and col_indices["lp"] < len(cells) else ""
                        opis_val = cells[col_indices["opis"]] if col_indices["opis"] is not None and col_indices["opis"] < len(cells) else ""
                        rez_val = cells[col_indices["rezultat"]] if col_indices["rezultat"] is not None and col_indices["rezultat"] < len(cells) else ""
                    except IndexError:
                        continue
                    
                    if lp_val or opis_val:
                        # Tworzymy płaski rekord na bazie aktualnego kontekstu wymagań i scenariusza
                        flat_row = context.copy()
                        flat_row["LP"] = lp_val
                        flat_row["Kroki testowe"] = opis_val
                        flat_row["Oczekiwany rezultat"] = rez_val
                        
                        all_rows.append(flat_row)
            
            else:
                # To nie tabela kroków – to tabela parametrów/wymagań (pary klucz-wartość)
                for row in table.rows:
                    cells = [c.text.strip() for c in row.cells]
                    if len(cells) >= 2:
                        key = cells[0].lower()
                        val = cells[1]
                        
                        if "moduł" in key or "modul" in key:
                            context["Moduł"] = val
                        elif "nr wymagania" in key or "numer wymagania" in key:
                            context["Pełny Nr wymagania"] = val
                        elif "opis wymagania" in key:
                            context["Opis wymagania"] = val
                        elif "cel" in key:
                            context["Cel"] = val
                        elif "warunki" in key:
                            context["Warunki wstępne"] = val

    # Budowanie końcowego DataFrame
    df = pd.DataFrame(all_rows)
    
    # Uzupełnienie brakujących kolumn wymaganych przez strukturę docelową
    for col in TARGET_COLUMNS:
        if col not in df.columns:
            df[col] = ""
            
    return df[TARGET_COLUMNS]

# --- SEKCJA INTERFEJSU STREAMLIT ---
uploaded_file = st.file_uploader("Wgraj plik scenariuszy Word (.docx)", type=["docx"])

if uploaded_file is not None:
    with st.spinner("Trwa głęboka analiza struktury dokumentu..."):
        try:
            df_result = parse_strict_uat_docx(uploaded_file)
            
            if not df_result.empty:
                st.success(f"Pomyślnie przetworzono dokument! Wykryto {len(df_result)} kroków testowych.")
                
                # Prezentacja danych w postaci KPI i Tabeli podglądu
                col1, col2 = st.columns(2)
                with col1:
                    st.metric("Liczba unikalnych wymagań", df_result["Pełny Nr wymagania"].nunique())
                with col2:
                    st.metric("Liczba scenariuszy testowych", df_result["Nr scenariusza"].nunique())

                st.subheader("👀 Podgląd spłaszczonej tabeli przed eksportem")
                st.dataframe(df_result, use_container_width=True)
                
                # Generowanie pliku Excel do pobrania
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    df_result.to_excel(writer, index=False, sheet_name='UAT_Scenariusze')
                buffer.seek(0)
                
                st.download_button(
                    label="📥 Pobierz gotowy arkusz Excel (.xlsx)",
                    data=buffer,
                    file_name="skonwertowane_scenariusze_uat.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.error("Struktura pliku różni się od oczekiwanego wzorca dokumentacji Comfortel. Nie udało się wyodrębnić kroków testowych.")
        except Exception as e:
            st.error(f"Wystąpił błąd krytyczny podczas parsowania pliku: {e}")
