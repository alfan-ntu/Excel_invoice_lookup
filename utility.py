#
# Subject: utility.py includes a few gadgets to support operation of this application
#   1. convert comma-separated currency string to it float type
#   2. ToDo's: debug-level controlled print
#
# Programmer: alfan-ntu
# Date: 2020/10/4
#
import logging


def initialization():
    # set filemode='w' to simply output log of the current run
    logging.basicConfig(filename="./log/excel_lookup.log", filemode='w', format='%(asctime)s %(levelname)s:%(message)s',
                        datefmt='%Y/%m/%d %I:%M:%S %p', level=logging.INFO)


def comma_separated_amount_to_float(amt_str):
    amt_no_comma = ""
    for i in range(len(amt_str)):
        if amt_str[i] != ",":
            amt_no_comma += amt_str[i]
    amount = float(amt_no_comma)
    return amount

