# This program process the rules of an excel file an return the needed code to put in Talend TMap expressions.

from openpyxl import load_workbook
from itertools import groupby
from collections import namedtuple

excel_rules_file_path = 'excel/regras_msp.xlsx'

wb = load_workbook(excel_rules_file_path) # get the workbook

ws = wb.active # get the first worksheet from the file

# -------------------------------------------------------------------------------------------------------------------

# This module gathers data from the excel rules file.
def dataToDict():
    # rules = [(<column_name>, [<column_values>])]
    # example: rules = {('CÒDIGO_DA_LINHA', ['E_01', 'E_02', ...]), ('Operação', [0, 0, 0, 0]), ...}
    rules = {} # a dictionary with {key:value} -> {<string>, <list>}
    for col in ws.iter_cols(max_row=19, min_col=2, max_col=6):
        for i, cell in enumerate(col):
            if i == 0:
                key = cell.value
                rules.setdefault(key, [])
            else:
                rules[key].append(cell.value)
    return rules

dict_rules = dataToDict()

# print(dict_rules)
# {'CÓDIGO DA LINHA': ['E_01', 'E_02', 'E_03', 'E_04', 'E_05', 'E_06', 'E_07', 'E_08', 'E_09', 'E_10', 'E_11', 'E_12', 'E_13', 'S_01', 'S_02', 'S_03', 'S_04', 'S_05'], 
# 'OPERAÇÃO': [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 'N/A', 1], 
# 'EMITENTE': [1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 0, 0, 0, 'N/A', 0]}

# -------------------------------------------------------------------------------------------------------------------

# This module applies the rules creating different sets for each rule

# Option1 Classes:
# applied_rules =  'operação': (0, ['E_01', 'E_02', 'E_03']), (1, ['E_04', 'E_05', 'E_06']), ...
# Class Rules:
    # Rules.name ex: operacao
    # Rules.rule ex: [0, 1, 'N/A']
    # Rules.result ex: (0, ['E_01', 'E_02', 'E_03']), (1, ['E_04', 'E_05', 'E_06']), ('N/A', ['S_01', 'S_02'])

# Option2: Named Tuples

# excel_rows = zip(*rules_values)  # zip(rules['Código da Linha'], rules['operação'], rules['emitente'], ...)
# print(rules_values)
# [('E_01', 0, 1), ('E_02', 0, 1), ('E_03', 0, 1), ('E_04', 0, 1), ('E_05', 0, 1), ('E_06', 0, 1), ('E_07', 0, 1), ('E_08', 0, 1), ('E_09', 0, 1), 
# ('E_10', 0, 1), ('E_11', 0, 1), ('E_12', 0, 1), ('E_13', 0, 1), ('S_01', 1, 0), ('S_02', 1, 0), ('S_03', 1, 0), ('S_04', 'N/A', 'N/A'), ('S_05', 1, 0)]
 
Rules = namedtuple('Rules', ['name', 'category', 'grouped_by_rules'])
rules_values = dict_rules.values()
rules_keys = list(dict_rules.keys())
rules = []
# Rules('OPERAÇÃO', [0, 1, 'N/A'], 
# { 0: ['S_01', 'S_02', 'S_03', 'S_05'], 1: ['E_01', 'E_02', 'E_03', 'E_04', 'E_05', 'E_06', 'E_07', 'E_08', 'E_09', 'E_10', 'E_11', 'E_12', 'E_13'], 'N/A': ['S_04']})

for i, column in enumerate(rules_values): # [(0, ["E_01", "E_02", "E_03" ...]), (1, [0, 0, 0, 1, 'N/A' ...]), (2, [1, 1, 1, 0, 'N/A'])]
    if i == 0:
        list_of_line_codes = column
    else:
        line_codes_and_rules = sorted(zip(list_of_line_codes, column), key= lambda x: str(x[1])) 
        # print(list(line_codes_and_rules)) # ("E_01", 0), ("E_02", 0)...
        categories = set()
        grouped_by_rules = {}
        for key, group in groupby(line_codes_and_rules, key= lambda x: x[1]): # the key is always the second element of the tuple
            # print("key: {0}, group: {1}".format(key, group))
            # prints:
            # -----------------------------------------------------------------
            # Grouped Rules:
            # key: 0, group: <itertools._grouper object at 0x000002D456D9E5C0>
            # key: 1, group: <itertools._grouper object at 0x000002D456D9EDA0>
            # key: N/A, group: <itertools._grouper object at 0x000002D456D9EFD0>
            # key: 0, group: <itertools._grouper object at 0x000002D455B2DCC0>
            # key: 1, group: <itertools._grouper object at 0x000002D456091D68>
            # key: N/A, group: <itertools._grouper object at 0x000002D455B2DCC0>
            # -----------------------------------------------------------------
            categories.add(key)
            grouped_by_rules[key] = [ pair[0] for pair in group]

        rules.append(Rules(rules_keys[i], categories, grouped_by_rules))

# print(grouped_by_rules)

print("=" * 40)
print("Testing named tuples")
for r in rules:
    print("=" * 40)
    print("Nome: {}".format(r.name))
    print("Categorias: {}".format(r.category))
    print("Regras agrupadas: {}".format(r.grouped_by_rules))


# -------------------------------------------------------------------------------------------------------------------

# This module write the code to use in Talend

