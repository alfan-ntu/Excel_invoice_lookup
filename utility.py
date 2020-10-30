#
# Subject: utility.py includes a few gadgets to support operation of this application
#   1. convert comma-separated currency string to it float type
#   2. ToDo's: debug-level controlled print
#
# Programmer: alfan-ntu
# Date: 2020/10/4
#
import xlrd
import logging
import constant


def initialization():
    # set filemode='w' to simply output log of the current run
    logging.basicConfig(filename="./log/excel_lookup.log", filemode='w', format='%(asctime)s %(levelname)s:%(message)s',
                        datefmt='%Y/%m/%d %I:%M:%S %p', level=logging.INFO)

#
# convert comma separated currency annotation to its floating number
#
def comma_separated_amount_to_float(amt_str):
    amt_no_comma = ""
    for i in range(len(amt_str)):
        if amt_str[i] != ",":
            amt_no_comma += amt_str[i]
    amount = float(amt_no_comma)
    return amount


#
# Looking for exchange rate from the remark column. SUPPOSE all currency exchange rates
# lead with "匯率"
#
def find_currency_exchange_rate(sourceRow):
    ex_rate = 1.0
    remark = sourceRow[constant.COL_INVOICE_REMARK].value
    idxEx = remark.find(constant.EXCHANGE_RATE_LEADING_CHRS)
    idxUsdAmt = remark.find(constant.USD_AMOUNT_CHRS)
    # it is an transaction of USD if both exchange_rate and USD sales amount found
    if idxEx >= 0 and idxUsdAmt >=0:
        ex_rate = float(remark[(idxEx+4):(idxUsdAmt-2)])

    return ex_rate


#
# Claim it is a USD transaction if "註記欄" includes both "匯率" and "美金未稅"
#
def is_source_a_usd_transaction(source_row):
    remark = source_row[constant.COL_INVOICE_REMARK].value
    idx_ex = remark.find(constant.EXCHANGE_RATE_LEADING_CHRS)
    idx_usd_amt = remark.find(constant.USD_AMOUNT_CHRS)
    if idx_ex >= 0 and idx_usd_amt >= 0:
        return True
    else:
        return False


#
# Claim the target transaction is a USD transaction if the column 'Currency Rate' > 1
#
def is_target_a_usd_transaction(targetRow):
    # pdb.set_trace()
    currency_exchange_rate = targetRow[constant.COL_GL_EXCHANGE_RATE].value
    ex_rate = float(currency_exchange_rate)
    if ex_rate > 1.0:
        return True
    else:
        return False


#
# Target records need to be further processed include
#   1. Account Description includes constant.TARGET_ACCOUNT_IN_GL (Accounts Receivable, actually)
#   2. Amount > 0; Amount <= 0 means Account Receivables received
#   3. Voucher Date is in the specified accounting period
#   ToDo's : needs to implement specified accounting period
#
def is_target_account_receivable(targetRow):
    targetCell = targetRow[constant.COL_GL_ACCOUNT_DESCRIPTION]
    account = targetCell.value
    idxAccount = account.find(constant.TARGET_ACCOUNT_IN_GL)
    amount = targetRow[constant.COL_GL_AMOUNT].value

    if idxAccount < 0:
        # print("Not an Account Receivable")
        return False
    if targetRow[constant.COL_GL_AMOUNT].ctype == xlrd.XL_CELL_EMPTY or \
            targetRow[constant.COL_GL_AMOUNT].ctype == xlrd.XL_CELL_BLANK:
        # print("Encountering Empty Amount Cell")
        return False
    # Looks like IFS output string of length 1, for which cell is not empty,
    # not blank, and float() fails to convert it to a floating number
    if targetRow[constant.COL_GL_AMOUNT].ctype == xlrd.XL_CELL_TEXT:
        if len(amount) == 1:
            # print("Not a valid amount")
            return False
    # amount of a record is negative when the "Account Receivable" is received
    if float(amount) < 0.0:
        # print("Account Receivable received")
        return False
    return True
