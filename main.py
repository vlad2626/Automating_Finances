import openpyxl

#global Variable
totalSpent=0
payments = 0
def load():
    global totalSpent
    global payments
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
# end of load
def calculations():
    print("total spent this month is ", totalSpent)
    print("total paid this month is ", payments)

load()
calculations()









