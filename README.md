# What it is?

This program process the rules of an excel file an return the needed code to put in Talend TMap expressions.


# How It Works ?

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

# What I used?

- Python 3.6
- openpyxl (excel library for python)