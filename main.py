from file_lib import Spreadsheet
from lib import TextFormElement, ButtonFormElement, SleepElement, Form


spreadsheet = Spreadsheet('sample.csv')


# To begin, we load the form on the supplied URL.
# This is just the starting point, so the URL will change as we navigate around.
form = Form('https://www.e-oscar-web.net/EntryController?trigger=Login')

# First we need to log in.
form.fill([
    form.select(id='companyId', val=''),
    form.select(id='userId', val=''),
    form.select(id='password', val=''),
    form.select(id='securityMsgAck1'),
    form.select(selector='.loginBtn')
])

#We need to get the correct AUD Panel

with form.frame('topFrame'):
    form.fill([
        form.select(id='DFProcessAUDanch')
    ])

#Once in we need to fill the AUD forms

with form.frame('leftFrame'):
    form.fill([
        form.select(id='DFCreateAUD')
    ])

with form.frame('mainFrame'):
        # We need the AUD Control number for reference
        print(form.select(id='headerInformation.audControlNumber').attribute("value"))
        form.fill([

            form.select(selector='.custom-combobox-input',val=''),
            form.select(selector='.custom-combobox-input', nth_child=1,val=''),
            form.select(selector='.custom-combobox-input', nth_child=3,val=''),
            form.select(id='headerInformation.audCorrectionIndicator',val='1:Update'),
            form.select(id='consumerInformation.lastName', val=''),
            form.select(id='consumerInformation.firstName', val=''),
            form.select(id='consumerInformation.generationCode',val=''),
            form.select(name='consumerInformation.ssnAreaNumber',val=''),
            form.select(name='consumerInformation.ssnGroupNumber',val=''),
            form.select(name='consumerInformation.ssnSerialNumber',val=''),
            form.select(id='consumerInformation.ecoaCode',val=''),
            form.select(id='consumerInformation.streetAddress',val=''),
            form.select(id='consumerInformation.city',val=''),
            form.select(id="consumerInformation.state",val=''),
            form.select(id='consumerInformation.zipCode1',val='')
        ])
with form.frame('leftFrame'):
            form.fill([
                form.select(id='DFAUDAccountInformationanch')
            ])
with form.frame('mainFrame'):
            form.fill([
                       form.select(id='accountInformation.accountStatusCode', val=''),
                       form.select(id='accountInformation.paymentRatingCode', val=''),
                       form.select(id='accountInformation.accountNumber', val=''),
                       form.select(id='accountInformation.portfolioTypeCode', val=''),
                       form.select(id='accountInformation.accountTypeCode', val=''),
                       form.select(name='accountInformation.dateOpenedMonth', val=''),
                       form.select(name='accountInformation.dateOpenedDay', val=''),
                       form.select(name='accountInformation.dateOpenedYear', val=''),
                       form.select(name='accountInformation.accountInformationMonth', val=''),
                       form.select(name='accountInformation.accountInformationDay', val=''),
                       form.select(name='accountInformation.accountInformationYear', val=''),
                       form.select(name='accountInformation.lastPaymentMonth', val=''),
                       form.select(name='accountInformation.lastPaymentDay', val=''),
                       form.select(name='accountInformation.lastPaymentYear', val=''),
                       form.select(name='accountInformation.dateClosedMonth', val=''),
                       form.select(name='accountInformation.dateClosedDay', val=''),
                       form.select(name='accountInformation.dateClosedYear', val=''),
                       form.select(name='accountInformation.fcra1stDelinquencyMonth', val=' '),
                       form.select(name='accountInformation.fcra1stDelinquencyDay', val=' '),
                       form.select(name='accountInformation.fcra1stDelinquencyYear', val=' '),
                       form.select(id='accountInformation.currentBalance', val=''),
                       form.select(id='accountInformation.amountPastDue', val=''),
                       form.select(id='accountInformation.highCreditOriginalAmount', val=''),
                       form.select(id='accountInformation.actualPaymentAmount', val=''),
                       form.select(id='Submit')
            ])
with form.frame('mainFrame'):
        form.fill([
            form.select(tag='input',nth_child=2)
        ])

print('Process Done!!!!')
