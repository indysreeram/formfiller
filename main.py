from file_lib import Spreadsheet
import argparse
import csv
import sys
from lib import TextFormElement, ButtonFormElement, SleepElement, Form, Notify
import time
ts = time.gmtime()
process_start_time = time.strftime("%Y-%m-%d %H:%M:%S", ts)




print('eOscar Woodbury Process Started at {} !!!!'.format(process_start_time))

#Lets check the arguments
body = 'Process Start time : ' + process_start_time
body = body + '\n'
logMsg = ''
parser = argparse.ArgumentParser()
parser.add_argument('--inputFile', nargs=1, required=True)
parser.add_argument('--outputFile', nargs=1, required=True)
parser.add_argument('--userId', nargs=1, required=True)
parser.add_argument('--password', nargs=1, required=True)
parser.add_argument('--company', nargs=1, required=True)
parser.add_argument('--headless',nargs=1,required=True)
args = parser.parse_args()


inputFile = str(args.inputFile[0])
outputFile = str(args.outputFile[0])
companyId = str(args.company[0])
userId = str(args.userId[0])
password = str(args.password[0])
headless = str(args.headless[0])


spreadsheet = Spreadsheet(inputFile)
# To begin, we load the form on the supplied URL.
# This is just the starting point, so the URL will change as we navigate around.
form = Form('https://www.e-oscar-web.net/EntryController?trigger=Login', headless)

# First we need to log in.
form.fill([
    form.select(id='companyId', val=companyId),
    form.select(id='userId', val=userId),
    form.select(id='password', val=password),
    form.select(id='securityMsgAck1'),
    form.select(selector='.loginBtn')
])

#We need to get the correct AUD Panel

with form.frame('topFrame'):
    form.fill([
        form.select(id='DFProcessAUDanch')
    ])

# For each row in the input csv we need to start filling the AUD forms
for row in spreadsheet.rows:
        row_dict = dict(zip(spreadsheet.header, row.cells))
        #Once in we need to fill the AUD forms

        with form.frame('leftFrame'):
              form.fill([
                  form.select(id='DFCreateAUD')
          ])
        with form.frame('mainFrame'):
                # We need the AUD Control number for reference
                print("Aud Control #",form.select(id='headerInformation.audControlNumber').attribute("value"), " for account Number " ,row_dict['account_number'])
                logMsg=''
                logMsg = "Aud Control #" + form.select(id='headerInformation.audControlNumber').attribute("value") + " for account Number "  +row_dict['account_number']
                print('logMsg = ',logMsg)
                row.set('aud_control_number', form.select(id='headerInformation.audControlNumber').attribute("value"))


                if len(row_dict['generation_code']) >= 4:
                     form.fill([
                          # form.select(selector='.custom-combobox-input', val=row_dict['equifax_subscriber_code']),
                          # form.select(selector='.custom-combobox-input', nth_child=1,val=row_dict['experian_subscriber_code']),
                          # form.select(selector='.custom-combobox-input', nth_child=3,val=row_dict['transunion_subscriber_code']),

                          form.select(selector='[list=EquifaxCodeList]', val=row_dict['equifax_subscriber_code']),
                          form.select(selector='[list=ExperianCodeList]', val=row_dict['experian_subscriber_code']),
                          form.select(selector='[list=TransunionCodeList]', val=row_dict['transunion_subscriber_code']),
                          form.select(id='headerInformation.audCorrectionIndicator',val=row_dict['aud_correction_indicator']),
                          form.select(id='consumerInformation.lastName', val=row_dict['last_name']),
                          form.select(id='consumerInformation.firstName', val=row_dict['first_name']),
                          form.select(id='consumerInformation.generationCode',val=row_dict['generation_code']),
                          form.select(name='consumerInformation.ssnAreaNumber',val=row_dict['ssn'][0:3]),
                          form.select(name='consumerInformation.ssnGroupNumber',val=row_dict['ssn'][3:5]),
                          form.select(name='consumerInformation.ssnSerialNumber',val=row_dict['ssn'][5:]),
                          form.select(id='consumerInformation.ecoaCode',val=row_dict['ecoa_code']),
                          form.select(id='consumerInformation.streetAddress',val=row_dict['street_address']),
                          form.select(id='consumerInformation.city', val=row_dict['city']),
                          form.select(id="consumerInformation.state", val=row_dict['state']),
                          form.select(id='consumerInformation.zipCode1', val=row_dict['zip_code'])
                            ])
                else:
                      form.fill([

                          # form.select(selector='.custom-combobox-input', val=row_dict['equifax_subscriber_code']),
                          # form.select(selector='.custom-combobox-input', nth_child=1,val=row_dict['experian_subscriber_code']),
                          # form.select(selector='.custom-combobox-input', nth_child=3,val=row_dict['transunion_subscriber_code']),

                          form.select(selector='[list=EquifaxCodeList]', val=row_dict['equifax_subscriber_code']),
                          form.select(selector='[list=ExperianCodeList]', val=row_dict['experian_subscriber_code']),
                          form.select(selector='[list=TransunionCodeList]', val=row_dict['transunion_subscriber_code']),

                          form.select(id='headerInformation.audCorrectionIndicator',val=row_dict['aud_correction_indicator']),
                          form.select(id='consumerInformation.lastName', val=row_dict['last_name']),
                          form.select(id='consumerInformation.firstName', val=row_dict['first_name']),
                          form.select(name='consumerInformation.ssnAreaNumber',val=row_dict['ssn'][0:3]),
                          form.select(name='consumerInformation.ssnGroupNumber',val=row_dict['ssn'][3:5]),
                          form.select(name='consumerInformation.ssnSerialNumber',val=row_dict['ssn'][5:]),
                          form.select(id='consumerInformation.ecoaCode',val=row_dict['ecoa_code']),
                          form.select(id='consumerInformation.streetAddress',val=row_dict['street_address']),
                          form.select(id='consumerInformation.city', val=row_dict['city']),
                          form.select(id="consumerInformation.state", val=row_dict['state']),
                          form.select(id='consumerInformation.zipCode1', val=row_dict['zip_code'])
                ])
        with form.frame('leftFrame'):
                    form.fill([
                        form.select(id='DFAUDAccountInformationanch')
                    ])
        with form.frame('mainFrame'):
                    form.fill([
                        form.select(id='accountInformation.accountStatusCode', val= row_dict['account_status']),
                        form.select(id='accountInformation.paymentRatingCode', val= row_dict['payment_rating']),
                        form.select(id='accountInformation.accountNumber', val= row_dict['account_number']),
                        form.select(id='accountInformation.portfolioTypeCode', val= row_dict['portfolio_type']),
                        form.select(id='accountInformation.accountTypeCode', val= row_dict['account_type']),
                        form.select(name='accountInformation.dateOpenedMonth', val= row_dict['date_opened'].split('/')[0]),
                        form.select(name='accountInformation.dateOpenedDay', val=row_dict['date_opened'].split('/')[1]),
                        form.select(name='accountInformation.dateOpenedYear', val=row_dict['date_opened'].split('/')[2]),
                        form.select(name='accountInformation.accountInformationMonth', val = row_dict['date_of_account_information'].split('/')[0]),
                        form.select(name='accountInformation.accountInformationDay',  val = row_dict['date_of_account_information'].split('/')[1]),
                        form.select(name='accountInformation.accountInformationYear', val = row_dict['date_of_account_information'].split('/')[2]),
                        form.select(name='accountInformation.lastPaymentMonth', val = row_dict['date_of_last_payment'].split('/')[0]),
                        form.select(name='accountInformation.lastPaymentDay', val = row_dict['date_of_last_payment'].split('/')[1]),
                        form.select(name='accountInformation.lastPaymentYear', val = row_dict['date_of_last_payment'].split('/')[2]),
                        form.select(name='accountInformation.dateClosedMonth', val=row_dict['date_closed'].split('/')[0]),
                        form.select(name='accountInformation.dateClosedDay', val=row_dict['date_closed'].split('/')[1]),
                        form.select(name='accountInformation.dateClosedYear',val=row_dict['date_closed'].split('/')[2]),
                        form.select(name='accountInformation.fcra1stDelinquencyMonth', val= row_dict['fcra_1_date_of_deliquency'].split('/')[0] if len(row_dict['fcra_1_date_of_deliquency']) >=6 else ' '),
                        form.select(name='accountInformation.fcra1stDelinquencyDay', val= row_dict['fcra_1_date_of_deliquency'].split('/')[1] if len(row_dict['fcra_1_date_of_deliquency']) >=6 in row_dict else ' '),
                        form.select(name='accountInformation.fcra1stDelinquencyYear', val= row_dict['fcra_1_date_of_deliquency'].split('/')[2] if len(row_dict['fcra_1_date_of_deliquency']) >=6 in row_dict else ' '),
                        form.select(id='accountInformation.currentBalance', val= row_dict['current_balance']),
                        form.select(id='accountInformation.amountPastDue', val= row_dict['current_balance']),
                        form.select(id='accountInformation.highCreditOriginalAmount', val=row_dict['high_credit_or_original_amount']),
                        form.select(id='accountInformation.actualPaymentAmount', val=row_dict['actual_payment']),
                        form.select(id='Submit')

                    ])
        with form.frame('mainFrame'):
                form.fill([
                    form.select(tag='input',nth_child=2)
                        ])
        body = body + '\n' +logMsg
#Once done we need to logout


with form.frame('topFrame'):
    form.fill([
        form.select(id='Logout')
    ])

# We write the results to new file

with open(outputFile,mode='w') as csv_file:
    csv_writer = csv.writer(csv_file, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)
    csv_writer.writerow(spreadsheet.header)
    for row in spreadsheet.rows:
        csv_writer.writerow(row.cells)
process_end_time = time.strftime("%Y-%m-%d %H:%M:%S", ts)


body =  body + '\n' + 'Process End time : ' + process_end_time
# notify = Notify('YOUR_SMTP_SERVER','YOUR_PASSWORD','FROM_ADDRESS','TO_ADDRESS','eOScar Woordbury Process',body)
print (body)
# notify.send_email()
print('eOscar Woodbury Process Done at {} !!!!'.format(process_end_time))


