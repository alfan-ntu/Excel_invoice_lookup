#
# Subject: Compare two Excel files which include invoice information and general ledger for monthly
#          accounting purpose
# Coder: alfan-ntu
# Created Date: 2020/10/2
# Revision:
#   1. 2020/10/2: v. 0.1 1st creation
#   2. 2020/10/30: v. 0.2
#           - rearrange utility functions by moving some widgets to utility.py
#           - modify Invoice Details to include a matched mark
#
# ToDo's:
#   1. Add an argument parser to accept source invoice records, target general ledger file, specified invoicing date
#   2. Mark record matching status on both source invoice records and target general ledger file
#
# Note:
#   1. xlrd can extract data from Excel files of format, .xls or .xlsx
#   2. xlwt can generate spreadsheet file compatible with .xls (MS Excel 97/2000/XP/200) BUT NOT .xlsx
#   3. openpyxl can read/write Excel 2010 xlsx/xlsm/xltx/xltm files, BUT NOT .xls
#   Both xlrd/xlwt/xlutils and openpyxl packages are required to process the input
#   invoice files and general ledger files since Invoice Details is .xls while General Ledger
#   is .xlsx. Pandas might be a flexible and more versatile alternative.
#
import sys
import openpyxl
import xlsxwriter
import pdb
import xlrd
from xlutils.copy import copy as xlutils_copy

import constant
import class_transaction
import utility
import logging
import class_opts

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
    ex_rate = utility.find_currency_exchange_rate(sourceRow)
    if ex_rate == 1.0:
        pass
    else:
        print(sourceRow[constant.COL_INVOICE_NO].value, " is a transaction in USD$, exchange rate:", str(ex_rate))
    return True, 20


def main(argv):
    # process argv and opts
    opts_args = class_opts.Opts(argv)
    # Initialize the execution
    utility.initialization()
    # Open source invoice details Excel file
    # invoiceLoc = "./invoice_Details_20200930.xls"
    invoiceLoc = opts_args.invoice_file
    sheetName = "Sheet0"
    sourceWb = xlrd.open_workbook(invoiceLoc, formatting_info=True)
    sourceWs = sourceWb.sheet_by_name(sheetName)
    sourceWb_temp = xlutils_copy(sourceWb)
    sourceWs_temp = sourceWb_temp.get_sheet(0)
    #
    # Open target general ledger Excel file
    # generalLedger = "./Voucher_Row_Analysis_20200930.xlsx"
    generalLedger = opts_args.ledger_file
    #
    # openpyxl to read target general ledger in order to read/modify/write .xlsx files
    #
    targetWb = openpyxl.load_workbook(generalLedger)
    # 0-based index, index of worksheet #1 is 0
    targetWs_name = targetWb.sheetnames[0]
    targetWs = targetWb[targetWs_name]
    #
    # Traverse the source invoice records
    # pdb.set_trace()
    sourceWs_temp.write(0,constant.COL_INVOICE_CHECKED, "發票配對")
    for js in range(1, sourceWs.nrows):
        invoice_status = sourceWs.cell_value(js, constant.COL_INVOICE_STATUS)
        if invoice_status == "作廢":
            sourceWs_temp.write(js, constant.COL_INVOICE_CHECKED, "作廢")
            continue
        invoice_number = sourceWs.cell_value(js, constant.COL_INVOICE_NO)
        buyer_name = sourceWs.cell_value(js, constant.COL_INVOICE_BUYER)
        invoice_date = sourceWs.cell_value(js, constant.COL_INVOICE_DATE)
        amount_nt_str = sourceWs.cell_value(js, constant.COL_INVOICE_TOTAL)
        invoice_amount_nt = utility.comma_separated_amount_to_float(amount_nt_str)
        # determine if it is a USD transaction, extract the exchange rate if
        # it is a USD transaction
        if utility.is_source_a_usd_transaction(sourceWs.row(js)):
            function_currency = constant.FUNCTION_CURRENCY_USD
            exchange_rate = utility.find_currency_exchange_rate(sourceWs.row(js))
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
        # traverse the target worksheet and identify the record correspondent to the
        # source transaction
        match_found = False

        for jt in range(2, targetWs.max_row+1):
            if not utility.is_target_account_receivable(targetWs[jt]):
                continue
            invoice_number = targetWs.cell(row=jt, column=constant.COL_GL_INVOICE_NO+1).value
            buyer_name = targetWs.cell(row=jt, column=constant.COL_GL_TEXT+1).value
            invoice_date = targetWs.cell(row=jt, column=constant.COL_GL_INVOICE_DATE+1).value
            invoice_amount_nt = targetWs.cell(row=jt, column=constant.COL_GL_AMOUNT+1).value
            if utility.is_target_a_usd_transaction(targetWs[jt]):
                function_currency = constant.FUNCTION_CURRENCY_USD
                exchange_rate = targetWs.cell(row=jt, column=constant.COL_GL_EXCHANGE_RATE+1).value
                invoice_amount_us = invoice_amount_nt / exchange_rate
            else:
                function_currency = constant.FUNCTION_CURRENCY_NTD
                exchange_rate = 1.0
                invoice_amount_us = 0.0
            source = constant.DATA_SOURCE_GENERAL_LEDGER
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
                sourceWs_temp.write(js, constant.COL_INVOICE_CHECKED, "是")

        if jt == targetWs.max_row and match_found is False:
            logging.info(">>>>>>>>>>>>>> 無法找到匹配交易紀錄 <<<<<<<<<<<<<<<, targetWs.max_row %s", targetWs.max_row)
            logging.info("==========================================================")
            sourceWs_temp.write(js, constant.COL_INVOICE_CHECKED, "否")

    sourceWb_temp.save(invoiceLoc)


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
    main(sys.argv[0:])
