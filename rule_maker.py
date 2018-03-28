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

How I named variables:
CODIGO_DA_LINHA, OPERAÇÃO, BLOCO # These are rules
E_01, E_02 # these are codes
(0, 1, 'N/A') or (C100, D100, C180) # these are categories


Step 1: Read and Process
dataToDict(): return dict_rules

Example:
{
    'CÓDIGO DA LINHA': ['E_01', 'E_02', 'E_03', 'S_01'], 
    'OPERAÇÃO': [0, 0, 1, 'N/A'], 
    'BLOCO': [C100, D100, C180, C180]
    }

Step 2: Applie rules
applie_rules(dict_rules): return rules_obj_list # a list containning each rule object.
A rule object have: .name, .category and .grouped_by_rules attributes
Example:

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
format_to_talend(rules): return fully_formated_rule

Example: BLOCO rule
row1.LINHA_APURACAO != null && 
row1.LINHA_APURACAO.equalsIgnoreCase("E_11") ? "D100" : ""

"""
import argparse
import re
from collections import namedtuple
from itertools import groupby
from os import mkdir

from openpyxl import load_workbook

def main():

    # rule_maker.py menu
    # https://docs.python.org/3.3/library/argparse.html?highlight=arg
    # TODO too many args :/
    parser = argparse.ArgumentParser(
        description='Read a excel file, process its rules and format the result for input in Talend TMap Expressions.'
    )
    group1 = parser.add_argument_group('Excel row and columns configuration.')
    group1.add_argument(
        'row_heading', type=int,
        help='numero da linha onde esta o cabecalho das regras'
    )
    group1.add_argument(
        'last_data_row', type=int,
        help='numero da ultima linha que ainda contem valores'
    )
    group1.add_argument(
        'first_column', type=int, metavar='first_col',
        help='number of the first column'
    )
    group1.add_argument(
        'last_column', type=int, metavar='last_col',
        help='number of the last column'
    )

    parser.add_argument(
        'directory_address', metavar='dir_addr',
        help="the name of the directory to use (creates it if it doesn't exist)",
        
    )
    parser.add_argument(
        'project_title', metavar='proj_title',
        help='the name of the current project to use as a version control inside each file created, and as a directory name'
    )
    args = parser.parse_args()
    print(args.row_heading)
    if args.directory_address:
        directory_address = args.directory_address
    else: # TODO Can I do this as a Function default (function name: write_to_file) ?
        directory_address = "regras_" + args.project_title + "/"

    dict_rules = dataToDict(
        row_heading=args.row_heading,
        last_data_row=args.last_data_row,
        first_column=args.first_column,
        last_column=args.last_column
    )

    rules_obj_list = applie_rules(dict_rules)

    for rule in rules_obj_list:
        print(rule.name)
        print("-"*len(rule.name))
        fully_formated_rule = format_to_talend(rule)
        # print(fully_formated_rule)
        write_to_file(rule.name, fully_formated_rule, directory_address=directory_address)


    # print(grouped_by_rules)

    # print("=" * 40)
    # print("Testing named tuples")

    # for r in applied_rules:
    #     print("=" * 40)
    #     print("Nome: {}".format(r.name))
    #     print("Categorias: {}".format(r.category))
    #     print("Regras agrupadas: {}".format(r.grouped_by_rules))


# -------------------------------------------------------------------------------------------------------------------

# This module gathers data from the excel rules file.
def dataToDict(row_heading, last_data_row, first_column, last_column):

    excel_rules_file_path = 'excel/regras_msp_ferreirav3.xlsx'
    wb = load_workbook(excel_rules_file_path) # get the workbook
    ws = wb.active # get the first worksheet from the file

    dict_rules = {} # a dictionary with {key:value} -> {<string>, <list>}

    for col in ws.iter_cols(min_row=row_heading, max_row=last_data_row, min_col=first_column, max_col=last_column):
        for i, cell in enumerate(col):
            if i == 0:
                key = cell.value
                dict_rules.setdefault(key, [])
            else:
                dict_rules[key].append(cell.value)
    return dict_rules

# def excel_configuration(worksheet):

#     class breakThroughLoops(Exception): pass
    
#     # find the the top and left limits
#     try:
#         for row in worksheet.rows:
#             for cell in row:
#                 m = re.search(r'\D_\d+', cell.value)
#                 if m:
#                     location = cell.colum + cell.row  # example 'B6'
#                     row_heading = int(cell.row) - 1 # 6-1 =5
#                     first_column = int(cell.col_idx)  # 2
#                     raise breakThroughLoops
#     except breakThroughLoops:
#         pass
    
#     # find the bottom limit
#     try:
#         for col in worksheet.iter_cols(min_row=cell.row, min_col=cell.col_idx, max_col=cell.col_idx):
#             for cell in col:
#                 n = re.search(r'\D_\d+', cell.value)
#                 if not n:
#                     last_data_row = int(cell.col_idx) -1
#                     raise breakThroughLoops
#     except breakThroughLoops:
#         pass

#     # find the right limit

#     return row_heading, last_data_row, first_column, last_column


# -------------------------------------------------------------------------------------------------------------------

# This module applies the rules creating different sets for each rule

def applie_rules(dict_rules): # what name should I use?

    Rules = namedtuple('Rules', ['name', 'category', 'grouped_by_rules'])
    rules_values = dict_rules.values() # I iterate over it
    rules_keys = list(dict_rules.keys()) # I call each element to name each Rule
    rules_obj_list = []
    
    for i, column in enumerate(rules_values): # [(0, ["E_01", "E_02", "E_03" ...]), (1, [0, 0, 0, 1, 'N/A' ...]), (2, [1, 1, 1, 0, 'N/A'])]
        if i == 0:
            list_of_line_codes = column
        elif ('INFORMADO NA BASE' in column):
            continue
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

            rules_obj_list.append(Rules(rules_keys[i], categories, grouped_by_rules))

    return rules_obj_list # return a list of named tuples

# -------------------------------------------------------------------------------------------------------------------

# This module format the expressions to use in Talend

def format_to_talend(rule_obj):

    fully_formated_rule = ""
        
    for category in rule_obj.category:
        initial_condition = "row1.LINHA_APURACAO != null && \n"
        list_of_codes = rule_obj.grouped_by_rules
        # for element in list_of_codes[category]:  ToDo
        
        if(len(list_of_codes[category]) == 1):
            #do something
            category_condition = "row1.LINHA_APURACAO.equalsIgnoreCase(\"{code}\") ? \"{rule}\" : \n".format(
                code=list_of_codes[category][0], rule=category
            )
            formated_category =  initial_condition + category_condition
            fully_formated_rule += formated_category
        else:
            #do other thing
            for i, code in enumerate(list_of_codes[category]): 
                if(i == 0):
                    category_condition = "(" + "row1.LINHA_APURACAO.equalsIgnoreCase(\"{code}\")".format(code=code)
                else: # len(list_of_codes) = 4  [0, 1, 2, 3]
                    category_condition += " || row1.LINHA_APURACAO.equalsIgnoreCase(\"{code}\")".format(code=code)
                    if( i % 2 !=0):
                        category_condition += '\n'
            
            category_condition = category_condition.rstrip('\n') + ")" + " ? \"{rule}\" : \n".format(rule=category)
            formated_category =  initial_condition + category_condition
            fully_formated_rule += formated_category

    fully_formated_rule += "\"\" "
    return fully_formated_rule

# -------------------------------------------------------------------------------------------------------------------

# This module write each fully_formated_rule into separeted .txt files

def write_to_file(rule_name, fully_formated_rule, directory_address, project_title='MSP'):

    try:
        mkdir(directory_address)
    except FileExistsError:
        pass

    file_address = directory_address + '/' + '{}.txt'.format(rule_name)
    try:
        # if file exists
        with open(file_address, "r+") as f:
            file_data = f.read()
            new_title = version_control(file_data)
            full_text = '\n' + new_title + '\n' + fully_formated_rule
            f.seek(0,0)
            f.write(full_text + '\n' + file_data)
    except FileNotFoundError:
        # if file does not exist
        with open(file_address, "w+") as f:
            full_text = '\n' + project_title + ' 1' + '\n' + fully_formated_rule
            f.write(full_text)
    finally:
        f.close()
        print("Formated rule for talend successfully created in {}".format(file_address))

def version_control(file_data):

    end = file_data.find('\n', 2)
    title_list = file_data[:end].split(" ")
    new_version = int(title_list[1]) + 1
    new_title = title_list[0].lstrip('\n') + ' ' + str(new_version)
    return new_title


main()