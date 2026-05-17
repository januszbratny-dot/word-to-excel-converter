import pandas as pd
import os

def przygotuj_scenariusze_uat(input_csv, output_xlsx):
    print("1. Wczytywanie danych z pliku CSV...")
    # Wczytujemy plik CSV (dostosuj separator, jeśli Twój plik używa innego niż przecinek)
    df = pd.read_csv(input_csv, sep=',')
    
    print("2. Naprawianie brakujących kroków testowych...")
    
    # --- POPRAWKA 1: Brakujący krok 1 dla WSF_BSS_OU_003 (TAKC_001) ---
    krok_1_data = {
        'Scenariusze testowe': 'Scenariusze testowe [SORT.BSS | Obsługa Umów]',
        'Moduł': 'SORT.BSS',
        'Pełny Nr wymagania': 'WSF_BSS_OU_003',
        'Nr scenariusza': 'TAKC_001',
        'Nazwa scenariusza': 'Konfiguracja pozycji słownikowej w kreatorze ofert',
        'LP': 1,
        'Kroki testowe': 'Użytkownik naciska na przycisk “Dodaj”. Wypełnia pola “Kod” oraz “Opis” i naciska na przycisk “Zapisz”',
        'Oczekiwany rezultat': 'Pozycja słownikowa została poprawnie dodana i zapisana w systemie.'
    }
    
    # Warunek logiczny wskazujący, gdzie wstawić wiersz (przed krokiem nr 2 tego scenariusza)
    idx_003 = df[(df['Pełny Nr wymagania'] == 'WSF_BSS_OU_003') & 
                 (df['Nr scenariusza'] == 'TAKC_001') & 
                 (df['LP'] == 2)].index
    
    if not idx_003.empty:
        target_idx = idx_003[0]
        # Tworzymy tymczasowy DataFrame z brakującym krokiem
        row_to_insert = pd.DataFrame([krok_1_data], columns=df.columns)
        # Składamy DataFrame na nowo rozcinając go w miejscu brakującego kroku
        df = pd.concat([df.iloc[:target_idx], row_to_insert, df.iloc[target_idx:]]).reset_index(drop=True)
        print(" -> Pomyślnie dodano brakujący Krok 1 do scenariusza WSF_BSS_OU_003 (TAKC_001)")

    # --- POPRAWKA 2: Brakujący krok 4 dla WSF_BSS_OU_004 (TAKC_001) ---
    krok_4_data = {
        'Scenariusze testowe': 'Scenariusze testowe [SORT.BSS | Obsługa Umów]',
        'Moduł': 'SORT.BSS',
        'Pełny Nr wymagania': 'WSF_BSS_OU_004',
        'Nr scenariusza': 'TAKC_001',
        'Nazwa scenariusza': 'Konfiguracja szablonu dla oferty',
        'LP': 4,
        'Kroki testowe': 'Użytkownik wprowadza nazwę np. “Podsumowanie oferty – klient indywidualny”, opcjonalnie wypełnia pole “opis”, następnie naciska na przycisk “Zapisz”',
        'Oczekiwany rezultat': 'Szablon oferty został poprawnie zapisany w konfiguracji systemu.'
    }
    
    idx_004 = df[(df['Pełny Nr wymagania'] == 'WSF_BSS_OU_004') & 
                 (df['Nr scenariusza'] == 'TAKC_001') & 
                 (df['LP'] == 5)].index
                 
    if not idx_004.empty:
        target_idx = idx_004[0]
        row_to_insert = pd.DataFrame([krok_4_data], columns=df.columns)
        df = pd.concat([df.iloc[:target_idx], row_to_insert, df.iloc[target_idx:]]).reset_index(drop=True)
        print(" -> Pomyślnie dodano brakujący Krok 4 do scenariusza WSF_BSS_OU_004 (TAKC_001)")

    print("3. Generowanie pliku Excel wraz z formatowaniem...")
    
    # Tworzymy ExcelWriter z silnikiem xlsxwriter
    writer = pd.ExcelWriter(output_xlsx, engine='xlsxwriter')
    df.to_excel(writer, sheet_name='Scenariusze UAT', index=False)
    
    workbook  = writer.book
    worksheet = writer.sheets['Scenariusze UAT']
    
    max_row = len(df) + 1  # Liczba wierszy danych + nagłówek
    max_col = len(df.columns) - 1
    
    # Zdefiniowanie stylów formatowania
    border_format = workbook.add_format({'border': 1})
    
    # Format ';;;' to ukryty format liczbowy. Ukrywa tekst, zachowując wartość w komórce.
    hide_duplicate_format = workbook.add_format({'num_format': ';;;'})
    
    # --- FORMATOWANIE 1: Dodanie obramowania do całego arkusza ---
    # Nakładamy ramki na wszystkie niepuste i puste komórki w obszarze danych
    worksheet.conditional_format(0, 0, max_row - 1, max_col,
                                 {'type': 'no_blanks', 'format': border_format})
    worksheet.conditional_format(0, 0, max_row - 1, max_col,
                                 {'type': 'blanks', 'format': border_format})
    
    # --- FORMATOWANIE 2: Ukrywanie duplikatów za pomocą formatowania warunkowego ---
    # Lista kolumn, w których treść powtarza się w pionie i powinna być ukryta (od indeksu 0 do 8)
    # Są to kolumny: Scenariusze testowe, Moduł, Nr wymagania, Opis, Zakres, Nr scenariusza, Nazwa, Cel, Warunki wstępne
    columns_to_hide_duplicates = [0, 1, 2, 3, 4, 5, 6, 7, 8]
    
    # Litery kolumn w Excelu odpowiadające indeksom (od A do I)
    column_letters = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I']
    
    for col_idx, col_letter in zip(columns_to_hide_duplicates, column_letters):
        # Reguła sprawdza czy bieżący wiersz (np. A2) jest równy wierszowi powyżej (A1)
        # Zaczynamy od wiersza 2 (indeks 1 w xlsxwriter) do samego końca tabeli
        formula = f'={col_letter}2={col_letter}1'
        
        worksheet.conditional_format(1, col_idx, max_row - 1, col_idx, {
            'type': 'formula',
            'criteria': formula,
            'format': hide_duplicate_format
        })
        
    # --- AUTODOPASOWANIE SZEROKOŚCI KOLUMN (Opcjonalnie dla czytelności) ---
    for i, col in enumerate(df.columns):
        max_len = max(df[col].astype(str).map(len).max(), len(col)) + 3
        # Ograniczamy szerokość kolumn, by bardzo długie opisy nie rozciągały ekranu bez końca
        worksheet.set_column(i, i, min(max_len, 40))
        
    # Zapisujemy plik
    writer.close()
    print(f"Sukces! Plik został zapisany jako: {output_xlsx}")

# --- URUCHOMIENIE SKRYPTU ---
if __name__ == "__main__":
    # Nazwa Twojego wejściowego pliku CSV wyeksportowanego z Excela
    wejsciowy_plik = "scenariusze_uat_odbior_final.xlsx - Scenariusze UAT.csv"
    wyjsciowy_plik = "Scenariusze_UAT_Sformatowane_Final.xlsx"
    
    if os.path.exists(wejsciowy_plik):
        przygotuj_scenariusze_uat(wejsciowy_plik, wyjsciowy_plik)
    else:
        print(f"Błąd: Nie znaleziono pliku źródłowego '{wejsciowy_plik}' w bieżącym katalogu.")
