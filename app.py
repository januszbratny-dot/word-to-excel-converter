import pandas as pd

# Zakładam, że masz już swój DataFrame o nazwie df
# df = ...

# Ustawienie engine na xlsxwriter
writer = pd.ExcelWriter('scenariusze_z_formatowaniem.xlsx', engine='xlsxwriter')
df.to_excel(writer, sheet_name='Scenariusze', index=False)

workbook  = writer.book
worksheet = writer.sheets['Scenariusze']

max_row = len(df) + 1  # Wiersze danych + nagłówek
max_col = len(df.columns) - 1

# 1. Dodanie obramowania dla całej tabeli
border_format = workbook.add_format({'border': 1})

# Aplikujemy obramowanie na wszystkie niepuste komórki
worksheet.conditional_format(0, 0, max_row - 1, max_col,
                             {'type': 'no_blanks', 'format': border_format})
worksheet.conditional_format(0, 0, max_row - 1, max_col,
                             {'type': 'blanks', 'format': border_format})

# 2. Formatowanie warunkowe ukrywające duplikaty
# Format ';;;' sprawia, że zawartość komórki jest niewidoczna dla oka
hide_format = workbook.add_format({'num_format': ';;;'})

# Przykład: Ukrywanie duplikatów dla kolumny A (indeks 0). 
# Zakładamy, że dane zaczynają się od wiersza 2 (indeks 1 w xlsxwriter).
# Formuła sprawdza, czy komórka jest równa tej powyżej (np. =A2=A1)
worksheet.conditional_format(1, 0, max_row - 1, 0, {
    'type': 'formula',
    'criteria': '=A2=A1',
    'format': hide_format
})

# Możesz powielić to dla innych kolumn, np. dla kolumny B (indeks 1):
# worksheet.conditional_format(1, 1, max_row - 1, 1, {
#     'type': 'formula',
#     'criteria': '=B2=B1',
#     'format': hide_format
# })

# Zapisanie pliku
writer.close()
