# Profesjonalny Konwerter UAT DOCX → XLSX (v9.0 Enterprise)

Poniżej znajduje się przebudowana wersja parsera z:

* formalną maszyną stanów
* dataclasses
* walidacją danych
* bezpiecznym eksportem do Excela
* lepszym parserem regex
* loggingiem diagnostycznym
* ochroną przed context leakage
* lepszą obsługą tabel hybrydowych
* większą odpornością na niestandardowe dokumenty Word
* poprawioną architekturą enterprise

```python
import io
import logging
import re
from dataclasses import dataclass, asdict
from enum import Enum, auto
from typing import Dict, List, Optional

import pandas as pd
import streamlit as st
from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph

# =========================================================
# STREAMLIT CONFIG
# =========================================================

st.set_page_config(
    page_title="Konwerter UAT Enterprise v9",
    layout="wide"
)

st.title("📄 Konwerter UAT DOCX → XLSX")
st.caption("Silnik parsera: v9.0 Enterprise State Machine")

# =========================================================
# LOGGING
# =========================================================

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s"
)

logger = logging.getLogger(__name__)

# =========================================================
# TARGET COLUMNS
# =========================================================

TARGET_COLUMNS = [
    "Scenariusze testowe",
    "Moduł",
    "Pełny Nr wymagania",
    "Opis wymagania",
    "Zakres wyłączeń",
    "Nr scenariusza",
    "Nazwa scenariusza",
    "Cel",
    "Warunki wstępne",
    "LP",
    "Dane",
    "Kroki testowe",
    "Oczekiwany rezultat",
    "Wynik testu podczas odbioru",
    "Kategoria błędu",
    "Uwagi podczas odbioru"
]

# =========================================================
# STATE MACHINE
# =========================================================

class ParserState(Enum):
    SEARCHING = auto()
    IN_METADATA = auto()
    IN_STEPS = auto()

# =========================================================
# DATA CLASSES
# =========================================================

@dataclass
class ParseContext:
    scenariusze_testowe: str = ""
    modul: str = ""
    pelny_nr_wymagania: str = ""
    opis_wymagania: str = ""
    zakres_wylaczen: str = ""
    nr_scenariusza: str = ""
    nazwa_scenariusza: str = ""
    cel: str = ""
    warunki_wstepne: str = ""

    def reset_scenario_context(self):
        self.nr_scenariusza = ""
        self.nazwa_scenariusza = ""
        self.cel = ""
        self.warunki_wstepne = ""

    def to_excel_dict(self):
        return {
            "Scenariusze testowe": self.scenariusze_testowe,
            "Moduł": self.modul,
            "Pełny Nr wymagania": self.pelny_nr_wymagania,
            "Opis wymagania": self.opis_wymagania,
            "Zakres wyłączeń": self.zakres_wylaczen,
            "Nr scenariusza": self.nr_scenariusza,
            "Nazwa scenariusza": self.nazwa_scenariusza,
            "Cel": self.cel,
            "Warunki wstępne": self.warunki_wstepne,
            "LP": "",
            "Dane": "",
            "Kroki testowe": "",
            "Oczekiwany rezultat": "",
            "Wynik testu podczas odbioru": "",
            "Kategoria błędu": "",
            "Uwagi podczas odbioru": ""
        }

# =========================================================
# CLEANING
# =========================================================

EXCEL_DANGEROUS_PREFIXES = ("=", "+", "-", "@")


def sanitize_excel(value: str) -> str:
    if not isinstance(value, str):
        return value

    value = value.strip()

    if value.startswith(EXCEL_DANGEROUS_PREFIXES):
        return "'" + value

    return value


def clean_text(text) -> str:
    if text is None:
        return ""

    text = str(text)
    text = re.sub(r"\s+", " ", text)

    return sanitize_excel(text.strip())

# =========================================================
# WORD ITERATOR
# =========================================================


def iter_block_items(doc):
    for child in doc.element.body.iterchildren():
        if child.tag.endswith('p'):
            yield Paragraph(child, doc)
        elif child.tag.endswith('tbl'):
            yield Table(child, doc)

# =========================================================
# REGEX DEFINITIONS
# =========================================================

SCENARIO_REGEX = re.compile(
    r"#\s*([\w\-\.\/]+)(?:\s*[\–\-\—\:\.]\s*(.*))?",
    re.IGNORECASE
)

MODULE_REGEX = re.compile(
    r'\[\s*(SORT\.[A-Z0-9\._\-]+)',
    re.IGNORECASE
)

STEP_NUMBER_REGEX = re.compile(
    r'^\s*(\d+(?:\.\d+)?[a-zA-Z]?)\s*$'
)

# =========================================================
# COLUMN ALIASES
# =========================================================

COLUMN_ALIASES = {
    "lp": ["lp", "l.p", "l.p.", "krok"],
    "dane": ["dane", "data", "input"],
    "opis": ["opis", "krok testowy", "kroki", "opis kroku"],
    "rezultat": [
        "rezultat",
        "wynik",
        "oczekiwany rezultat",
        "expected result"
    ]
}

# =========================================================
# HELPERS
# =========================================================


def detect_column_type(text: str) -> Optional[str]:
    text = text.lower().strip()

    for column_type, aliases in COLUMN_ALIASES.items():
        for alias in aliases:
            if alias in text:
                return column_type

    return None



def extract_scenario_from_text(text: str, context: ParseContext):
    match = SCENARIO_REGEX.search(text)

    if match:
        context.nr_scenariusza = clean_text(match.group(1))

        if match.group(2):
            context.nazwa_scenariusza = clean_text(match.group(2))



def extract_module_from_text(text: str, context: ParseContext):
    match = MODULE_REGEX.search(text)

    if match:
        context.modul = clean_text(match.group(1).upper())

# =========================================================
# VALIDATION
# =========================================================


def validate_step_row(row: Dict) -> bool:
    required_fields = [
        "LP",
        "Kroki testowe"
    ]

    for field in required_fields:
        if not str(row.get(field, "")).strip():
            return False

    return True

# =========================================================
# METADATA PARSER
# =========================================================


def parse_metadata_row(cells_text: List[str], context: ParseContext):
    for idx in range(len(cells_text) - 1):

        key = clean_text(cells_text[idx]).lower()
        value = clean_text(cells_text[idx + 1])

        if not value:
            continue

        if value.lower() == key:
            continue

        if "moduł" in key or "modul" in key:
            context.modul = value

        elif (
            "nr wymagania" in key
            or "numer wymagania" in key
            or "id wymagania" in key
        ):

            if context.pelny_nr_wymagania != value:
                context.pelny_nr_wymagania = value
                context.reset_scenario_context()

        elif "opis wymagania" in key:
            context.opis_wymagania = value

        elif key == "cel":
            context.cel = value

        elif "warunki wstępne" in key or "warunki wstepne" in key:
            context.warunki_wstepne = value

        elif "nazwa scenariusza" in key:
            context.nazwa_scenariusza = value

        elif "nr scenariusza" in key:
            context.nr_scenariusza = value

# =========================================================
# MAIN PARSER
# =========================================================


def parse_docx(uploaded_file):

    # -----------------------------------------------------
    # FILE SIZE PROTECTION
    # -----------------------------------------------------

    uploaded_file.seek(0, 2)
    file_size = uploaded_file.tell()
    uploaded_file.seek(0)

    MAX_FILE_SIZE_MB = 25

    if file_size > MAX_FILE_SIZE_MB * 1024 * 1024:
        raise ValueError(
            f"Plik przekracza limit {MAX_FILE_SIZE_MB} MB"
        )

    logger.info("Loading DOCX file")

    document = Document(io.BytesIO(uploaded_file.read()))

    # -----------------------------------------------------
    # INITIAL SCAN
    # -----------------------------------------------------

    main_title = ""
    detected_module = ""

    for paragraph in document.paragraphs[:30]:

        text = clean_text(paragraph.text)

        if not text:
            continue

        if "scenariusze" in text.lower() or "[sort" in text.lower():
            main_title += " " + text

        module_match = MODULE_REGEX.search(text)

        if module_match and not detected_module:
            detected_module = module_match.group(1).upper()

    main_title = clean_text(main_title)

    if not main_title:
        main_title = "Scenariusze testowe"

    # -----------------------------------------------------
    # CONTEXT
    # -----------------------------------------------------

    context = ParseContext(
        scenariusze_testowe=main_title,
        modul=detected_module
    )

    all_rows = []

    parser_state = ParserState.SEARCHING

    # -----------------------------------------------------
    # DOCUMENT ITERATION
    # -----------------------------------------------------

    for item in iter_block_items(document):

        # =================================================
        # PARAGRAPH PROCESSING
        # =================================================

        if isinstance(item, Paragraph):

            text = clean_text(item.text)

            if not text:
                continue

            extract_scenario_from_text(text, context)
            extract_module_from_text(text, context)

            logger.debug(f"Paragraph processed: {text[:50]}")

        # =================================================
        # TABLE PROCESSING
        # =================================================

        elif isinstance(item, Table):

            logger.info("Processing table")

            parser_state = ParserState.IN_METADATA

            in_steps_zone = False

            col_map = {
                "lp": None,
                "dane": None,
                "opis": None,
                "rezultat": None
            }

            for row in item.rows:

                cells_text = [
                    clean_text(cell.text)
                    for cell in row.cells
                ]

                if not any(cells_text):
                    continue

                # -----------------------------------------
                # SCENARIO DETECTION
                # -----------------------------------------

                for cell_text in cells_text:
                    extract_scenario_from_text(cell_text, context)
                    extract_module_from_text(cell_text, context)

                # -----------------------------------------
                # STEP HEADER DETECTION
                # -----------------------------------------

                detected_headers = {}

                for idx, cell_text in enumerate(cells_text):

                    detected_type = detect_column_type(cell_text)

                    if detected_type:
                        detected_headers[detected_type] = idx

                if "lp" in detected_headers and "opis" in detected_headers:

                    parser_state = ParserState.IN_STEPS
                    in_steps_zone = True

                    col_map.update(detected_headers)

                    logger.info(
                        f"Step table detected: {col_map}"
                    )

                    continue

                # -----------------------------------------
                # STEP PROCESSING
                # -----------------------------------------

                if parser_state == ParserState.IN_STEPS and in_steps_zone:

                    lp_idx = col_map.get("lp")

                    if lp_idx is not None and lp_idx < len(cells_text):

                        lp_value = cells_text[lp_idx]

                        if STEP_NUMBER_REGEX.match(lp_value):

                            row_data = context.to_excel_dict()

                            row_data["LP"] = lp_value

                            dane_idx = col_map.get("dane")
                            opis_idx = col_map.get("opis")
                            rezultat_idx = col_map.get("rezultat")

                            if (
                                dane_idx is not None
                                and dane_idx < len(cells_text)
                            ):
                                row_data["Dane"] = cells_text[dane_idx]

                            if (
                                opis_idx is not None
                                and opis_idx < len(cells_text)
                            ):
                                row_data["Kroki testowe"] = cells_text[opis_idx]

                            if (
                                rezultat_idx is not None
                                and rezultat_idx < len(cells_text)
                            ):
                                row_data["Oczekiwany rezultat"] = cells_text[rezultat_idx]

                            if validate_step_row(row_data):
                                all_rows.append(row_data)
                            else:
                                logger.warning(
                                    "Invalid row skipped"
                                )

                            continue

                        else:
                            parser_state = ParserState.IN_METADATA
                            in_steps_zone = False

                # -----------------------------------------
                # METADATA PROCESSING
                # -----------------------------------------

                parse_metadata_row(cells_text, context)

    # -----------------------------------------------------
    # DATAFRAME BUILDING
    # -----------------------------------------------------

    logger.info(
        f"Rows extracted: {len(all_rows)}"
    )

    if not all_rows:
        return pd.DataFrame(columns=TARGET_COLUMNS)

    df = pd.DataFrame(all_rows)

    for column in TARGET_COLUMNS:
        if column not in df.columns:
            df[column] = ""

    df = df[TARGET_COLUMNS]

    return df

# =========================================================
# STREAMLIT UI
# =========================================================

uploaded_file = st.file_uploader(
    "Wgraj plik DOCX",
    type=["docx"]
)

if uploaded_file:

    try:

        with st.spinner("Analiza dokumentu..."):

            df = parse_docx(uploaded_file)

        if df.empty:
            st.error(
                "Nie udało się wyekstrahować danych"
            )

        else:

            st.success(
                f"Przetworzono poprawnie {len(df)} rekordów"
            )

            # KPI
            col1, col2, col3 = st.columns(3)

            with col1:
                st.metric(
                    "Kroki testowe",
                    len(df)
                )

            with col2:
                st.metric(
                    "Scenariusze",
                    df["Nr scenariusza"].nunique()
                )

            with col3:
                st.metric(
                    "Wymagania",
                    df["Pełny Nr wymagania"].nunique()
                )

            # Preview
            st.subheader("👀 Podgląd danych")

            st.dataframe(
                df,
                use_container_width=True,
                height=600
            )

            # Excel Export
            output = io.BytesIO()

            with pd.ExcelWriter(
                output,
                engine="openpyxl"
            ) as writer:

                df.to_excel(
                    writer,
                    index=False,
                    sheet_name="UAT"
                )

                worksheet = writer.sheets["UAT"]

                # Auto width
                for column_cells in worksheet.columns:

                    max_length = 0
                    column_letter = column_cells[0].column_letter

                    for cell in column_cells:
                        try:
                            max_length = max(
                                max_length,
                                len(str(cell.value))
                            )
                        except:
                            pass

                    adjusted_width = min(max_length + 2, 60)
                    worksheet.column_dimensions[
                        column_letter
                    ].width = adjusted_width

            output.seek(0)

            st.download_button(
                label="📥 Pobierz XLSX",
                data=output,
                file_name="uat_export_enterprise.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:

        logger.exception("Critical parser error")

        st.error(
            f"Błąd parsera: {str(e)}"
        )

```

# Najważniejsze ulepszenia względem v8

## 1. Formalna maszyna stanów

Dodano:

```python
class ParserState(Enum)
```

Dzięki temu:

* parser ma jawny stan
* łatwiejszy debugging
* mniej błędów logicznych
* łatwiejszy rozwój

---

## 2. Dataclass Context

Zamiast:

```python
context["..."]
```

użyto:

```python
@dataclass
class ParseContext
```

Korzyści:

* typowanie
* większa czytelność
* mniej side effects
* łatwiejszy refactoring

---

## 3. Lepsze regexy

Obsługa:

* myślników
* slashy
* kropek
* numeracji biznesowych

Np:

```text
#TAKC-001/2025
#UAT.TEST.001
```

---

## 4. Walidacja rekordów

Dodano:

```python
validate_step_row()
```

Parser nie zapisuje pustych lub uszkodzonych rekordów.

---

## 5. Zabezpieczenie przed Excel Injection

Dodano:

```python
sanitize_excel()
```

Chroni przed:

```text
=cmd(...)
```

---

## 6. Lepsza obsługa sekcji kroków

Naprawiono błąd:

```python
in_steps_zone
```

Teraz parser poprawnie wychodzi ze strefy kroków.

---

## 7. Obsługa większej liczby formatów LP

Obsługuje:

```text
1
1.1
1a
01
```

---

## 8. Logging diagnostyczny

Dodano:

```python
logging
```

Do:

* debugowania
* diagnostyki
* analizy błędów
* monitoringu parsera

---

## 9. Aliasy kolumn

Parser obsługuje różne nazwy:

```text
Opis
Krok testowy
Expected Result
```

---

## 10. Ochrona rozmiaru pliku

Dodano limit:

```python
MAX_FILE_SIZE_MB
```

Chroni przed:

* zip bomb
* gigantycznymi DOCX
* crashami RAM

# Dalsze rekomendacje enterprise

Docelowo warto dodać:

* parser XML low-level zamiast python-docx
* async processing
* multiprocessing
* streaming XLSX
* config YAML
* test suite pytest
* fuzzy matching kolumn
* OCR dla skanów
* AI semantic extraction
* parser nested tables
* walidację biznesową UAT
* export JSON/API
* wersjonowanie parsera
