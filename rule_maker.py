# This program process the rules of an excel file an return the needed code to put in Talend TMap expressions.
# Python Path:
# "python.pythonPath": "C:/Users/vitor.rabelo/virtual_enviroments/ruletomatic/Scripts/python.exe"
"""
There is a column called 'CODIGO_DA_LINHA' (code_of_line) in wich will be applied a set of rules. This set of rules correspond
to the other columns.
Exemple:
CODIGO_DA_LINHA || OPERAÇÃO || BLOCO
    E_01              0         C100
    E_02              0         D100
    E_03              1         C180
    S_01            'N/A'       C180

The first column is the reference, the others are the rules. This module will read, process, applie the rules and format
it to use in Talend.
Step 1: Read and Process
dataToDict(): return dict_rules
{
    'CÓDIGO DA LINHA': ['E_01', 'E_02', 'E_03', 'S_01'], 
    'OPERAÇÃO': [0, 0, 1, 'N/A'], 
    'BLOCO': [C100, D100, C180, C180]
    }

Step 2: Applie rules
applie_rules(dict_rules): return rules (# Is it a good name?)
Nome: OPERAÇÃO
Categorias: {0, 1, 'N/A'}
Regras agrupadas: {
    0: ['E_01', 'E_02'],
    1: ['E_03'], 
    'N/A': ['S_01']
    }
========================================
Nome: BLOCO
Categorias: {'D100', 'C180', 'C100'}
Regras agrupadas: {
    'C100': ['E_01'], 
    'C180': ['E_03', 'S_01'], 
    'D100': ['E_02']
    }
========================================

Step 3: Format the results


"""

from openpyxl import load_workbook
from itertools import groupby
from collections import namedtuple

excel_rules_file_path = 'excel/regras_msp.xlsx'

wb = load_workbook(excel_rules_file_path) # get the workbook

ws = wb.active # get the first worksheet from the file

def main():

    dict_rules = dataToDict()

    applied_rules = applie_rules(dict_rules) 

    # print(grouped_by_rules)

    print("=" * 40)
    print("Testing named tuples")

    for r in applied_rules:
        print("=" * 40)
        print("Nome: {}".format(r.name))
        print("Categorias: {}".format(r.category))
        print("Regras agrupadas: {}".format(r.grouped_by_rules))


# -------------------------------------------------------------------------------------------------------------------

# This module gathers data from the excel rules file.
def dataToDict():
    dict_rules = {} # a dictionary with {key:value} -> {<string>, <list>}
    for col in ws.iter_cols(max_row=19, min_col=2, max_col=6):
        for i, cell in enumerate(col):
            if i == 0:
                key = cell.value
                dict_rules.setdefault(key, [])
            else:
                dict_rules[key].append(cell.value)
    return dict_rules

# -------------------------------------------------------------------------------------------------------------------

# This module applies the rules creating different sets for each rule

def applie_rules(dict_rules): # what name should I use?

    Rules = namedtuple('Rules', ['name', 'category', 'grouped_by_rules'])
    rules_values = dict_rules.values() # I iterate over it
    rules_keys = list(dict_rules.keys()) # I call each element to name each Rule
    rules = []
    
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

    return rules # what name should I use?

# -------------------------------------------------------------------------------------------------------------------

# This module write the code to use in Talend



# -------------------------------------------------------------------------------------------------------------------

main()