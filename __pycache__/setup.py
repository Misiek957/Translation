import pandas as pd
selection = ['Psychology', 'Historia', 'Sociology']
for select in selection:
    table = pd.read_excel("working_wordlist.xlsx", sheet_name=select)
    original_table = table
    NaN = ' '
    for i in range(2, 50):
        # Create 50 new fields for possible translations.
        original_table['1' + 3 * '0' + str(i)] = NaN
        original_table['-1' + 3 * '0' + str(i)] = NaN
        original_table['0' + 3 * '0' + str(i)] = NaN
    original_table.to_excel("working_wordlist_" + select + ".xlsx", sheet_name=select)

