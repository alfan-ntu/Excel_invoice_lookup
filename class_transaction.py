#
# Widget set
#
from datetime import datetime
import constant
import logging
import pdb


class Transaction:
    def __init__(self, invoice_number, buyer_name, invoice_date, invoice_amount_nt,
                 invoice_amount_us, function_currency, exchange_rate,
                 transaction_data_source):
        # string: invoice_number, buyer_name, invoice_date
        # float: invoice_amount_NT, invoice_amount_US, exchange_rate
        self.invoice_number = invoice_number
        self.buyer_name = buyer_name
        self.invoice_date = invoice_date
        self.invoice_amount_NT = invoice_amount_nt
        self.function_currency = function_currency
        self.exchange_rate = exchange_rate
        if self.function_currency == constant.FUNCTION_CURRENCY_USD:
            self.invoice_amount_US = invoice_amount_nt / exchange_rate
        else:
            self.invoice_amount_US = invoice_amount_us

        self.source = transaction_data_source

    def display_transaction(self):
        logging.info("\t發票號碼: %s", self.invoice_number)
        logging.info("\t買方名稱: %s", self.buyer_name)
        if self.source == constant.DATA_SOURCE_INVOICE_DETAIL:
            date_object = datetime.strptime(self.invoice_date, "%Y/%m/%d")
        else:
            date_object = datetime.strptime(self.invoice_date, "%m/%d/%Y")
        logging.info("\t發票日期: %s", date_object)
        if self.function_currency == constant.FUNCTION_CURRENCY_USD:
            # print("交易類型: 美金交易/交易匯率@", str(self.exchange_rate))
            logging.info("\t交易類型: 美金交易/交易匯率@ %s", str(self.exchange_rate))
            # print("發票金額:", self.invoice_amount_NT, "of type:", type(self.invoice_amount_NT))
            logging.info("\t發票金額: %s", self.invoice_amount_NT)
            # print("美金/台幣匯率", self.exchange_rate)
            logging.info("\t美金/台幣匯率: %s", self.exchange_rate)
            # print("交易美金金額:", "%.2f" % (self.invoice_amount_NT / self.exchange_rate))
            logging.info("\t交易美金金額: %s", "%.2f" % (self.invoice_amount_NT / self.exchange_rate))
        else:
            # print("交易類型: 台幣交易")
            logging.info("\t交易類型: 台幣交易")
            # print("發票金額:", self.invoice_amount_NT, "of type:", type(self.invoice_amount_NT))
            logging.info("\t發票金額: %s", self.invoice_amount_NT)

    #
    # match_transaction() : matches the calling transaction object to the target transaction object
    # Match Criteria:
    #   1. buyer's name in source transaction partially matches buyer's name in the target transaction
    #   2. difference between the amount in source transaction and the amount in target transaction is
    #      within 1% of the amount in the source transaction
    #   3. (optional) invoice date in source transaction is the same as that in the target transaction
    #
    def match_transaction(self, target_transaction):
        buyer_in_source = self.buyer_name[0:4]
        buyer_in_target = target_transaction.buyer_name
        if buyer_in_target.find(buyer_in_source) < 0:
            return False
        if self.function_currency == constant.FUNCTION_CURRENCY_NTD:
            # function_currency_in_source = self.function_currency
            amount_in_source = self.invoice_amount_NT
            amount_diff_threshold = amount_in_source * constant.AMOUNT_DIFF_THRESHOLD_RATIO
            if type(target_transaction.invoice_amount_NT) is str:
                print("target_transaction.invoice_amount_NT is of type", type(target_transaction.invoice_amount_NT))
                print("target_transaction.invoice_amount_NT:", target_transaction.invoice_amount_NT)
                print("target_transaction.invoice_number:", target_transaction.invoice_number)
                amount_diff = amount_diff_threshold + 1
            else:
                amount_diff = target_transaction.invoice_amount_NT - amount_in_source
            if abs(amount_diff) > amount_diff_threshold:
                return False
            return True
        else:
            # FUNCTION_CURRENCY_USD
            amount_in_source = self.invoice_amount_US
            amount_diff_threshold = amount_in_source * constant.AMOUNT_DIFF_THRESHOLD_RATIO
            amount_diff = target_transaction.invoice_amount_US - amount_in_source
            if abs(amount_diff) > amount_diff_threshold:
                return False
            return True



