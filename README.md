# gnumeric
Python User Defined Functions to Extend Gnumeric

Installing the gnumeric-plugins-extra package allows you to write user defined functions in Python and use them natively in gnumeric just as any other built-in function.

You can also just modify and import the functions and use them in your program as you desire with the exception of the gnumeric specific functions.

All packages must be installed before hand. If there are errors in the python module because of packages or otherwise, Gnumeric will throw a general error which is: Function implementation not available.

You need to install:
  gnumeric-plugins-extra (e.g. using apt in Ubuntu)
  rdoclient (install using pip. This is only required for the true random functions. If you do not need these functions you can discard this step and delete the import and random functions before use, otherwise gnumeric will throw an error for all functions.) 
  
  Included files:
    plugin.xml                xml file required for function and module name definitions
    plugin_functions.py       python functions and helper functions to be used with gnumeric
    iban_validator.gnumeric   spreadsheet file showing the use of the python iban validator function
    ghost.gnumeric            spreadsheet file showing the ghost add proof of concept function
    
Included functions:
    py_ghost_add                  a proof of concept function, description inside above code
    py_ghost_control              ghost addition helper function
    py_true_random                get a true random number from Random.org
    py_signed_true_random         get a signed true random number from Random.org
    py_series_sum                 experimental function, just ignore this one
    py_iban_validate              validate an IBAN
    py_iban_bank_code             get IBAN bank code for supported countries only
    py_iban_branch_code           get IBAN branch code for supported countries only
    py_iban_account_number        get IBAN account number for supported countries only
    
