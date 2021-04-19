#
# Subject: This class processes the input arguments for executing the program
# Coder: alfan-ntu
# Created Date: 2020/10/2
# Revision:
#   1. 2020/10/2: v. 0.1 1st creation
#   2. 2020/11/1: v. 0.2
#           - added a new option for the external sales output file
#
import getopt
import sys
from datetime import datetime


class Opts:
    # Class Opts stores arguments to run this program, which includes
    #   1. invoice_details: Excel file where invoice details are stored in
    #   2. target_general_ledger: target Excel file where the general ledger
    #   3. invoice_date_start: starting date of the range of invoice date
    #   4. invoice_date_end: end date of the range of invoice date
    #
    def __init__(self, argv):
        # string: invoice_details (i:, --invoice), general_ledger (l:, --ledger)
        # date: invoice_date_start (b:, --begin), invoice_date_end (e:, --end)
        # switch: help (h, --help)
        self.invoice_file = ""
        self.ledger_file = ""
        self.sales_file = ""
        self.begin_date = ""
        self.end_date = ""
        try:
            opts, args = getopt.getopt(argv[1:], "hi:l:b:e:o:",
                                       ["help", "invoice=", "ledger=", "output=", "begin=", "end="])
        except getopt.GetoptError:
            print("Invalid command syntax...")
            print_help_message(argv[0])
            sys.exit()
        for opt, arg in opts:
            if opt in ("-h", "--help"):
                print_help_message(argv[0])
                sys.exit()
            elif opt in ("-i", "--invoice"):
                self.invoice_file = arg
            elif opt in ("-l", "--ledger"):
                self.ledger_file = arg
            elif opt in ("-b", "--begin"):
                self.begin_date = arg
            elif opt in ("-e", "--end"):
                self.end_date = arg
            elif opt in ("-o", "--output"):
                self.sales_file = arg
        if self.sales_file == "":
            self.sales_file = "External_Sales.xlsx"
        self.date_sanity_check()

    def date_sanity_check(self):
        if self.begin_date == "" and self.end_date == "":
            return True
        if self.begin_date != "":
            try:
                self.begin_date = datetime.strptime(self.begin_date, "%Y%m%d")
            except ValueError:
                print("Wrong starting date format")
                sys.exit()
        if self.end_date != "":
            try:
                self.end_date = datetime.strptime(self.end_date, "%Y%m%d")
            except ValueError:
                print("Wrong end date format")
                sys.exit()
        if self.begin_date != "" and self.end_date == "":
            self.end_date = datetime.today()
            print("End date is set to ", self.end_date.strftime("%Y/%m/%d"))
        if self.end_date != "" and self.begin_date == "":
            self.begin_date = datetime(2000, 1, 1)
            print("Beginning date is set to ", self.begin_date.strftime("%Y/%m/%d"))
        if self.end_date < self.begin_date:
            print("End date is earlier than starting date....")
            sys.exit()
        return True


def print_help_message(command):
    print("Syntax: ", command, " -i [invoice] -l [ledger] -o <output> -b <start date> -e <end date>")
    print("\t-i (--invoice): Invoice file name <mandatory>")
    print("\t-l (--ledger): General ledger file name <mandatory>")
    print("\t-b (--begin): Beginning invoicing date: yyyymmdd <optional>")
    print("\t-e (--end): End invoicing date: yyyymmdd <optional>")
    print("\t-h (--help): Print this help menu")
