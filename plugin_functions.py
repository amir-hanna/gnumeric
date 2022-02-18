# Script is written in Python 2.x to work with my version of Gnumeric
from xml.etree.ElementTree import tostring
from Gnumeric import GnumericError, GnumericErrorVALUE
import Gnumeric
import string
from rdoclient import RandomOrgClient


# HELPER FUNCTIONS not available in Gnumeric

def func_cell_value(obj):
    '@HELPER FUNCTION returns cellValue in obj'

    cell_value = None
    try:
        cell = Gnumeric.functions['CELL']
        cell_value = cell('value', obj)
    except TypeError:
        raise GnumericError, GnumericErrorVALUE
    except NameError:
        raise GnumericError, GnumericErrorNAME
    else:
        return cell_value

def func_cell_ref(obj):
    '@HELPER FUNCTION returns cell_ref (address) as text in obj'

    cell_ref = None
    try:
        cell = Gnumeric.functions['CELL']
        cell_ref = cell('address', obj)
    except TypeError:
        raise GnumericError, GnumericErrorVALUE
    except NameError:
        raise GnumericError, GnumericErrorNAME
    else:
        return cell_ref.replace('$', '')

def func_cell_sheet(obj):
    '@HELPER FUNCTION returns name of the sheet in obj'

    cell_sheet = None
    try:
        cell = Gnumeric.functions['CELL']
        cell_sheet = cell('sheetname', obj)
    except TypeError:
        raise GnumericError, GnumericErrorVALUE
    except NameError:
        raise GnumericError, GnumericErrorNAME
    else:
        return cell_sheet

def func_cell_col(obj):
    '@HELPER FUNCTION returns number of the col in obj'

    cell_col = None
    try:
        cell = Gnumeric.functions['CELL']
        cell_col = cell('col', obj)
    except TypeError:
        raise GnumericError, GnumericErrorVALUE
    except NameError:
        raise GnumericError, GnumericErrorNAME
    else:
        return int(cell_col)

def func_cell_row(obj):
    '@HELPER FUNCTION returns number of the row in obj'

    cell_row = None
    try:
        cell = Gnumeric.functions['CELL']
        cell_row = cell('row', obj)
    except TypeError:
        raise GnumericError, GnumericErrorVALUE
    except NameError:
        raise GnumericError, GnumericErrorNAME
    else:
        return int(cell_row)

def func_cell_contents(obj):
    '@HELPER FUNCTION returns contents in obj'

    cell_contents = None
    try:
        cell = Gnumeric.functions['CELL']
        cell_contents = cell('contents', obj)
    except TypeError:
        raise GnumericError, GnumericErrorVALUE
    except NameError:
        raise GnumericError, GnumericErrorNAME
    else:
        return cell_contents

def func_set_cell_value(obj, str_val, col_offset=0, row_offset=0):
    c = func_get_cell_object(obj, col_offset, row_offset)
    c.set_text(str_val)

def func_get_cell_object(obj, col_offset=0, row_offset=0):
    wb = Gnumeric.workbooks()[0]
    s = wb.sheets()[0]
    c = s.cell_fetch(func_cell_col(obj)-1 + int(col_offset), func_cell_row(obj)-1 + int(row_offset))
    return c

def get_entered_text(obj, col_offset=0, row_offset=0):
    c = func_get_cell_object(obj, col_offset, row_offset)
    return c.get_entered_text()

def ghost_control (cell1, cell2, str_concat, value_cell, str_value_command):
    '@FUNCTION=PY_GHOST_CONTROL\n'\
    '@DESCRIPTION=Adds two numbers together.\n'\
    '@SYNTAX=py_ghost_control(cell1, cell2, str_concat, value_cell, str_value_command)\n'\
    '@EXAMPLES=An example\n'\
    'second line\n\n'\
    '@SEEALSO='

    # if input changed, set the formula in the value cell
    if str_concat != str(func_cell_value(cell1)).strip() + " " + str(func_cell_value(cell2)).strip():
        func_set_cell_value(value_cell, str_value_command)

    return "Control Cell; DO NOT DELETE"



# Add two numbers together and do not recalculate on sheet open. Recalculate only when values change
def ghost_add (cell1, cell2, value_cell, control_cell):
    '@FUNCTION=PY_GHOST_ADD\n'\
    '@DESCRIPTION=Adds two numbers together.\n'\
    '@SYNTAX=py_ghost_add(cell1, cell2, value_cell, control_cell)\n'\
    '@EXAMPLES=An example\n'\
    'second line\n\n'\
    '@SEEALSO='

    str_cell1 = str(func_cell_value(cell1)).strip()
    str_cell2 = str(func_cell_value(cell2)).strip()

    # get the current formula in the value cell
    the_formula = get_entered_text(value_cell)

    # leaving formula in and exiting before doing any real calculations to give time to copy the formula
    if get_entered_text(control_cell) == "":
        return "Copy formula, then populate any data in control cell to start."

    # set the formula in the control cell
    func_set_cell_value(control_cell, "=py_ghost_control(" + func_cell_ref(cell1) + "; " + func_cell_ref(cell2) + "; \"" + str_cell1 + " " + str_cell2 + "\"; " + func_cell_ref(value_cell) + "; \"" + the_formula + "\")")


    # remove formula from value cell
    func_set_cell_value(value_cell, "")

    if str_cell1 == "" or str_cell1 == "None" or str_cell2 == "" or str_cell2 == "None":
        return

    return float(str_cell1) + float(str_cell2)

def iban_bank_code (cell1):
    '@FUNCTION=PY_IBAN_BANK_CODE\n'\
    '@DESCRIPTION=Get Bank Code From IBAN.\n'\
    '@SYNTAX=py_iban_bank_code(IBAN cell reference)\n'\
    '@EXAMPLES=An example\n'\
    'second line\n\n'\
    '@SEEALSO='

    if iban_validate(cell1) == 'INVALID': return 'INVALID'

    str_cell1 = str(func_cell_value(cell1)).strip()
    country_code = str_cell1[:2].lower()

    if country_code == 'eg': 
        return str_cell1[4:8]
    elif country_code == 'it':
        return str_cell1[5:10]
    else:
        return 'Not Supported'

def iban_branch_code (cell1):
    '@FUNCTION=PY_IBAN_BRANCH_CODE\n'\
    '@DESCRIPTION=Get Branch Code From IBAN.\n'\
    '@SYNTAX=py_iban_branch_code(IBAN cell reference)\n'\
    '@EXAMPLES=An example\n'\
    'second line\n\n'\
    '@SEEALSO='

    if iban_validate(cell1) == 'INVALID': return 'INVALID'

    str_cell1 = str(func_cell_value(cell1)).strip()
    country_code = str_cell1[:2].lower()

    if country_code == 'eg': 
        return str_cell1[8:12]
    elif country_code == 'it':
        return str_cell1[10:15]
    else:
        return 'Not Supported'

def iban_account_number (cell1):
    '@FUNCTION=PY_IBAN_ACCOUNT_NUMBER\n'\
    '@DESCRIPTION=Get Account Number from IBAN.\n'\
    '@SYNTAX=py_iban_account_number(IBAN cell reference)\n'\
    '@EXAMPLES=An example\n'\
    'second line\n\n'\
    '@SEEALSO='

    if iban_validate(cell1) == 'INVALID': return 'INVALID'

    str_cell1 = str(func_cell_value(cell1)).strip()
    country_code = str_cell1[:2].lower()

    if country_code == 'eg': 
        return str_cell1[12:]
    elif country_code == 'it':
        return str_cell1[15:]
    else:
        return 'Not Supported'

# validate iban
def iban_validate (cell1):
    '@FUNCTION=PY_IBAN_VALIDATE\n'\
    '@DESCRIPTION=Validates an IBAN.\n'\
    '@SYNTAX=py_iban_validate(cell1)\n'\
    '@EXAMPLES=An example\n'\
    'second line\n\n'\
    '@SEEALSO='

    str_cell1 = str(func_cell_value(cell1)).strip()
    
    country_code = str_cell1[:2].lower()
    iban_length = len(str_cell1)
    if (country_code == 'eg' and iban_length != 29) or (country_code == 'it' and iban_length != 27): 
        return 'INVALID'        

    lst = [ iban_convert(x) for x in str_cell1[4:] + str_cell1[:4] ]
    
    # replace long by int for Python 3.x
    if long(''.join(lst)) % 97 == 1:
        return 'VALID'
    return 'INVALID'

def iban_convert(x):
    if x.isalpha():
        return (str)(ord(x.lower()) - 87)
    return x

# get true random number from Random.org
def true_random (start_num, end_num):
    '@FUNCTION=PY_TRUE_RANDOM\n'\
    '@DESCRIPTION=Generate true random number from Random.org.\n'\
    '@SYNTAX=py_true_random(start_num, end_num)\n'\
    '@EXAMPLES=An example\n'\
    'second line\n\n'\
    '@SEEALSO='

    r = RandomOrgClient("Enter Your API Key Here From Random.ORG")
    result = r.generate_integers(1, start_num, end_num)
    return result[0]


# get signed true random number from Random.org
def signed_true_random (start_num, end_num):
    '@FUNCTION=PY_SIGNED_TRUE_RANDOM\n'\
    '@DESCRIPTION=Generate signed true random number from Random.org.\n'\
    '@SYNTAX=py_signed_true_random(start_num, end_num)\n'\
    '@EXAMPLES=An example\n'\
    'second line\n\n'\
    '@SEEALSO='

    r = RandomOrgClient("Enter Your API Key Here From Random.ORG")
    result = r.generate_signed_integers(1, start_num, end_num)

    if r.verify_signature(result["random"], result["signature"]) == True:
        return result["random"]["data"][0]
    else:
        return "error"

# Sum a hypothetical mathematical series
def func_series_sum (first_value, current_value, first_position, current_position, last_position):
    '@FUNCTION=PY_SERIES_SUM\n'\
    '@DESCRIPTION=Sum a hypothetical mathematical series by first calculating hypothetical increase.\n'\
    '@SYNTAX=py_series_sum (first_value, current_value, first_position, current_position, last_position)\n'\
    '@EXAMPLES=An example\n'\
    'second line\n\n'\
    '@SEEALSO='

    hypothetical_step = (current_value - first_value) / (current_position - first_position)
    number_of_items = last_position - first_position + 1
    last_value = first_value + (number_of_items - 1) * hypothetical_step
    series_sum = number_of_items * (first_value + last_value) / 2
    return series_sum



# Translate the python functions to gnumeric functions and register them
amir_functions = {
    'py_ghost_add': ('rrrr', 'cell1, cell2, value_cell, control_cell', ghost_add),
    'py_ghost_control': ('rrsrs', 'cell1, cell2, str_concat, value_cell, str_value_command', ghost_control),
    'py_true_random': ('ff', 'start_num, end_num', true_random),
    'py_signed_true_random': ('ff', 'start_num, end_num', signed_true_random),
    'py_series_sum': ('fffff', 'first_value, current_value, first_position, current_position, last_position', func_series_sum),
    'py_iban_validate': ('r', 'iban_cell_reference', iban_validate),
    'py_iban_bank_code': ('r', 'iban_cell_reference', iban_bank_code),
    'py_iban_branch_code': ('r', 'iban_cell_reference', iban_branch_code),
    'py_iban_account_number': ('r', 'iban_cell_reference', iban_account_number)
}
