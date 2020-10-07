import getopt
import sys


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
        self.begin_date = ""
        self.end_date = ""
        try:
            opts, args = getopt.getopt(argv[1:], "hi:l:b:e:",
                                       ["help", "invoice=", "ledger=", "begin=", "end="])
        except getopt.GetoptError:
            print("Syntax: \n\t", argv[0], " -i <invoice> -l <ledger> -b <start date> -e <end date>")
            sys.exit(2)
        for opt, arg in opts:
            if opt in ("-h", "--help"):
                print("Syntax: \n\t", argv[0], " -i <invoice> -l <ledger> -b <start date> -e <end date>")
                sys.exit()
            elif opt in ("-i", "--invoice"):
                self.invoice_file = arg
            elif opt in ("-l", "--ledger"):
                self.ledger_file = arg
            elif opt in ("-b", "--begin"):
                self.begin_date = arg
            elif opt in ("-e", "--end"):
                self.end_date = arg
        if self.invoice_file == "" or self.ledger_file == "":
            print("Invoice file or ledger file is mandatory")
            sys.exit()