import random
import sys
import time
from datetime import datetime

import pandas as pd
from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QFileDialog, QMessageBox
from random import randint
from datetime import date, timedelta
from ui_populateExcel import Ui_MainWindow


def exitBTN():
    sys.exit()


class appWindow(QtWidgets.QMainWindow, Ui_MainWindow):
    def __init__(self, parent=None):
        super(appWindow, self).__init__(parent)
        self.output_df = None
        self.dateList = []
        self.debitLedgerAmount_list = []

        self.setupUi(self)

        self.exitBTN.clicked.connect(exitBTN)
        self.generateBTN.clicked.connect(self.generateBTN_clicked)
        self.saveBTN.clicked.connect(self.saveBTN_clicked)

    def randomList(self, transactionCNT, totalAmount, amountFrom, amountTo):
        self.debitLedgerAmount_list = [0] * transactionCNT
        for i in range(totalAmount):
            self.debitLedgerAmount_list[randint(amountFrom, amountTo) % transactionCNT] += 1
        return self.debitLedgerAmount_list

    def generate_randDates(self, start, end, n):
        start = datetime.strptime(start, '%d/%m/%Y')
        end = datetime.strptime(end, '%d/%m/%Y')
        dates_bet = end - start
        total_days = dates_bet.days

        for idx in range(n):
            random.seed()
            randay = random.randrange(total_days)
            x = start + timedelta(days=randay)
            self.dateList.append(x)


        self.dateList = [date_obj.strftime('%d/%m/%Y') for date_obj in self.dateList]
        return self.dateList

    def get_TotalAmount(self):
        return int(self.amountTotal.text())

    def get_fromAmount(self):
        return int(self.amountFrom.text())

    def get_toAmount(self):
        return int(self.amountTo.text())

    def get_fromDate(self):
        return self.dateFrom.date().toString("dd/MM/yyyy")

    def get_toDate(self):
        return self.dateTo.date().toString("dd/MM/yyyy")

    def getTransactions(self):
        x = self.transactionCNT.value()
        return int(x)

    def getLedgerName(self):
        return self.ledgerName.text()

    def getDebitLedger(self):
        return self.ledgerDebit.text()

    def generateBTN_clicked(self):
        valid = self.inputsValidation()

        if valid:
            self.debitLedgerAmount_list = self.randomList(self.getTransactions(), self.get_TotalAmount(),
                                                          self.get_fromAmount(), self.get_toAmount())

            self.dateList = self.generate_randDates(self.get_fromDate(), self.get_toDate(), self.getTransactions())

            self.output_df = pd.DataFrame({'Date':self.dateList, 'Supplier Inv Date': self.dateList,
                                           'Debit Ledger 1 Amount':self.debitLedgerAmount_list})
            self.output_df.sort_values(by='Date', inplace=True)
            self.output_df['Voucher Type'] = self.output_df.apply(lambda _: 'Purchase', axis=1)
            self.output_df['IS Invoice'] = self.output_df.apply(lambda _: 'No', axis=1)
            self.output_df['Supplier Inv No'] = self.output_df.apply(lambda _: 'NA', axis=1)
            self.output_df['Credit / Party Ledger'] = self.output_df.apply(lambda _: self.getLedgerName(), axis=1)
            self.output_df['New_ID'] = range(0, len(self.output_df))
            self.output_df['Voucher No'] = self.output_df.apply(lambda _: self.getLedgerName()[:2], axis=1)
            self.output_df['Voucher No'] = self.output_df['Voucher No'].str.upper() + 'PUR' + self.output_df['New_ID'].astype(str)
            self.output_df['Debit Ledger 1'] = self.output_df.apply(lambda _: self.getDebitLedger(), axis=1)

            empty_cols = ['Address 1', 'Address 2', 'Address 3', 'Address 4', 'State', 'Place of Supply', 'VAT Tin No',
                          'CST No', 'Service Tax No', 'GSTIN', 'GST Registration Type']

            for col in empty_cols:
                self.output_df[col] = self.output_df.apply(lambda _: '', axis=1)

            self.output_df = self.output_df[['Date', 'Voucher No', 'Voucher Type', 'IS Invoice', 'Supplier Inv No',
                                             'Supplier Inv Date', 'Credit / Party Ledger', 'Address 1', 'Address 2',
                                             'Address 3', 'Address 4', 'State', 'Place of Supply', 'VAT Tin No',
                                             'CST No', 'Service Tax No', 'GSTIN', 'GST Registration Type', 'Debit Ledger 1',
                                             'Debit Ledger 1 Amount']]

            QMessageBox.about(self, 'Done', 'Execution Done Successfully. You may now save your file')
            self.generateBTN.setEnabled(False)

    def saveBTN_clicked(self):
        response = QFileDialog.getSaveFileName(caption='Save Consolidated Sheet',
                                               directory=f'{self.getLedgerName()} -- Output.xlsx',
                                               filter="Excel (*.xlsx *.xls *.csv)")

        if response[0] == "":
            QMessageBox.about(self, 'Caution', "Please specify save location!")
        else:
            writer = pd.ExcelWriter(response[0], engine='xlsxwriter')
            self.output_df.to_excel(writer, sheet_name='Output', startrow=0, startcol=0, index=False)
            writer.save()
            QMessageBox.about(self, 'Done', 'Save Done Successfully. you may now Exit the application.')

    def inputsValidation(self):
        if self.getTransactions() <= 0:
            QMessageBox.about(self, 'Caution', "Please specify number of transactions!")
        else:
            return True

        if self.get_TotalAmount() == 0:
            QMessageBox.about(self, 'Caution', "Please enter Total amount!")
        else:
            return True

        if type(self.get_TotalAmount()) != int:
            QMessageBox.about(self, 'Caution', "Please enter Total amount in numbers!")
        else:
            return True

        if type(self.get_TotalAmount()) != int:
            QMessageBox.about(self, 'Caution', "Please enter Total amount in numbers!")
        else:
            return True

        if type(self.get_fromAmount()) != int:
            QMessageBox.about(self, 'Caution', "Please enter Start amount in numbers!")
        else:
            return True

        if type(self.get_toAmount()) != int:
            QMessageBox.about(self, 'Caution', "Please enter End amount in numbers!")
        else:
            return True

        if self.get_toAmount() < self.get_fromAmount():
            QMessageBox.about(self, 'Caution', "End amount must be greater than Start amount!")
        else:
            return True


class Manager:
    def __init__(self):
        # Creating App Window
        self.appWindow = appWindow()

        # Start the program
        self.appWindow.show()


#####################
#        MAIN       #
#####################
if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    manager = Manager()
    sys.exit(app.exec_())
