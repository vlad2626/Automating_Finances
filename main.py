import pprint

import openpyxl

#global Variable
totalSpent=0
payments = 0
balance = 0
my_formatter = "{0:.2f}"
TOTALS =\
    {
        'Spent': my_formatter.format(totalSpent),
        'Paid': my_formatter.format(payments),
        'Balance': my_formatter.format(balance),
        'Month': 0
    }


def main():
    load()
    calculations()
    print(balance)
    writeFile()


def load():
    global totalSpent
    global payments
    global balance
    wb= openpyxl.load_workbook('Wellsfargo.xlsx')
    sheet = wb['CreditCard4']
    list(sheet.columns)[1]  # select the whole colum
    for cellObj in list(sheet.columns)[1]:
        value = cellObj.value
        #print(value)
        if value <0:
           totalSpent += value
        else:
            payments += value

    balance += totalSpent + payments
    print (balance)
# end of load
def calculations():

    TOTALS['Spent'] = my_formatter.format(totalSpent)
    TOTALS['Paid'] = my_formatter.format(payments)
    TOTALS['Balance'] = my_formatter.format(balance)
    TOTALS['Month'] = 2

    print (TOTALS)





#writes the output to a text file.
def writeFile():

    print('writing to a file')
    resultFile=open('totals.txt','a')
    resultFile.write('\nTotals =' + pprint.pformat(TOTALS))
    resultFile.close()
    print('Done. ')



main()







