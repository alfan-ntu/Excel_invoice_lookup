#
# Subject: Compare two Excel files which include invoice information and general ledger for monthly
#          accounting purpose
# Coder: alfan-ntu
# Created Date: 2020/10/2
# Revision:
#   1. 2020/10/2: 1st creation
#
# ToDo's:
#   1. Add an argument parser to accept source invoice records, target general ledger file, specified invoicing date
#   2. Add info logging feature
#
import xlsxwriter
import pdb
import xlrd
import constant
import class_transaction
import utility
import logging

#
# Function: match_row(sourceRow, targetWs)
# Subject: check if any record found in targetWs that
#   - buyerName,
#   - invoice status(開立、作廢),
#   - invoice date
#   - invoice amount
#   in sourceRow match the record in targetWs
# Input :
#       - sourceRow: the whole row sequence in the source invoice records Excel
#       - targetWs: the target worksheet which target general ledger stores in
# Output :
#       - a tuple
#
def match_row(sourceRow, targetWs):
    # pdb.set_trace()
    invoice_status = sourceRow[constant.COL_INVOICE_STATUS].value
    matchRow_in_targetWs = 0
    if invoice_status == "作廢":
        # print(sourceRow[6], "/", invoice_status)
        # print(sourceRow[constant.COL_INVOICE_NO].value, " is void")
        # return immediately if the invoice is void
        return False, matchRow_in_targetWs
    #
    # Continue traversing target worksheet only if the invoice in the sourceRow is a valid one
    # print(sourceRow[constant.COL_INVOICE_NO].value, " is valid")
    # for jt in (1, targetWs.nrows):
    ex_rate = find_currency_exchange_rate(sourceRow)
    if ex_rate == 1.0:
        pass
    else:
        print(sourceRow[constant.COL_INVOICE_NO].value, " is a transaction in USD$, exchange rate:", str(ex_rate))
    return True, 20


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
#   1. Account Description includes constant.TARGET_ACCOUNT_IN_GL
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
    if float(amount) < 0.0:
        # print("Account Receivable received")
        return False

    return True


def main():
    # Initialize the execution
    utility.initialization()
    # Open source invoice details Excel file
    invoiceLoc = "./invoice_Details_20200930.xls"
    sheetName = "Sheet0"
    sourceWb = xlrd.open_workbook(invoiceLoc)
    sourceWs = sourceWb.sheet_by_name(sheetName)

    # Open target general ledger Excel file
    # generalLedger = "./Voucher_Row_Analysis_20200930.xlsx"
    generalLedger = "./Voucher_Row_Analysis.xlsx"
    targetWb = xlrd.open_workbook(generalLedger)
    targetWs = targetWb.sheet_by_index(0)
    #
    # print(sourceWs.cell_value(0, 0))
    # for i in range(sourceWs.ncols):
    #     print(sourceWs.cell_value(0, i))
    #
    # Traverse the source invoice records
    #
    for js in range(1, sourceWs.nrows):
        invoice_status = sourceWs.cell_value(js, constant.COL_INVOICE_STATUS)
        if invoice_status == "作廢":
            continue
        invoice_number = sourceWs.cell_value(js, constant.COL_INVOICE_NO)
        buyer_name = sourceWs.cell_value(js, constant.COL_INVOICE_BUYER)
        invoice_date = sourceWs.cell_value(js, constant.COL_INVOICE_DATE)
        amount_nt_str = sourceWs.cell_value(js, constant.COL_INVOICE_TOTAL)
        invoice_amount_nt = utility.comma_separated_amount_to_float(amount_nt_str)
        # determine if it is a USD transaction, extract the exchange rate if
        # it is a USD transaction
        if is_source_a_usd_transaction(sourceWs.row(js)):
            function_currency = constant.FUNCTION_CURRENCY_USD
            exchange_rate = find_currency_exchange_rate(sourceWs.row(js))
        else:
            function_currency = constant.FUNCTION_CURRENCY_NTD
            exchange_rate = 1.00

        source = constant.DATA_SOURCE_INVOICE_DETAIL
        source_transaction = class_transaction.Transaction(invoice_number,
                                                          buyer_name,
                                                          invoice_date,
                                                          invoice_amount_nt,
                                                          0.0,
                                                          function_currency,
                                                          exchange_rate,
                                                          source)
        # call the class method to display the object contents
        source_transaction.display_transaction()
        # logging.info("----------------------------------------------------------")
        # traverse the target worksheet and identify the record correspondent to the
        # source transaction
        match_found = False
        for jt in range(1, targetWs.nrows):
            if not is_target_account_receivable(targetWs.row(jt)):
                continue
            invoice_number = targetWs.cell_value(jt, constant.COL_GL_INVOICE_NO)
            buyer_name = targetWs.cell_value(jt, constant.COL_GL_TEXT)
            invoice_date = targetWs.cell_value(jt, constant.COL_GL_INVOICE_DATE)
            invoice_amount_nt = targetWs.cell_value(jt, constant.COL_GL_AMOUNT)
            if is_target_a_usd_transaction(targetWs.row(jt)):
                function_currency = constant.FUNCTION_CURRENCY_USD
                exchange_rate = targetWs.cell_value(jt, constant.COL_GL_EXCHANGE_RATE)
                invoice_amount_us = invoice_amount_nt / exchange_rate
            else:
                function_currency = constant.FUNCTION_CURRENCY_NTD
                exchange_rate = 1.0
                invoice_amount_us = 0.0

            source = constant.DATA_SOURCE_GENERAL_LEDGER
        #     construct target Transaction object
            target_transaction = class_transaction.Transaction(invoice_number,
                                                               buyer_name,
                                                               invoice_date,
                                                               invoice_amount_nt,
                                                               invoice_amount_us,
                                                               function_currency,
                                                               exchange_rate,
                                                               source)
            if source_transaction.match_transaction(target_transaction):
                match_found = True
                logging.info(">>>>>>>>>>>>>> 找到匹配交易紀錄 <<<<<<<<<<<<<<<")
                target_transaction.display_transaction()
                logging.info("==========================================================")
            if match_found is False and jt == (targetWs.nrows-1):
                logging.info(">>>>>>>>>>>>>> 無法找到匹配交易紀錄 <<<<<<<<<<<<<<<")
                logging.info("==========================================================")


        # input("Type anything to continue")


def generate_excel(spread_sheet):
    workbook = xlsxwriter.Workbook("Customs_bill_records.xlsx")
    worksheet = workbook.add_worksheet("2nd_sheet")

    tax_bill = ""
    customs_bill = ""
    tax_ID = ""
    tax_amount = ""
    row = 0
    col = 0
    for tax_bill, customs_bill, tax_ID, tax_amount in spread_sheet:
        worksheet.write(row, col, tax_bill)
        worksheet.write(row, col+1, customs_bill)
        worksheet.write(row, col+2, tax_ID)
        worksheet.write(row, col+3, tax_amount)
        row += 1

    workbook.close()

def get_input_file():
    input_file_name = "201903.txt"
    hInputFile = open(input_file_name, "r", encoding="utf-8")
    return hInputFile

if __name__ == "__main__":
    main()
