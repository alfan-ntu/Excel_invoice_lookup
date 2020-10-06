#
# Subject: utility.py includes a few gadgets to support operation of this application
#   1. convert comma-separated currency string to it float type
#   2. ToDo's: debug-level controlled print
#
# Programmer: Al Fan@yapro.com.tw
# Date: 2020/10/4
#


def comma_separated_amount_to_float(amt_str):
    amt_no_comma = ""
    for i in range(len(amt_str)):
        if amt_str[i] != ",":
            amt_no_comma += amt_str[i]
    amount = float(amt_no_comma)
    return amount

