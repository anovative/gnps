from __future__ import print_function
import sys
# appending a path
import lib.auth as auth
import lib.services as services
from pathlib import Path

# import specific modules
import pandas as pd
import csv
#output folder
_outputFolder = Path("output")
_outputFileName = 'registration.data_check_results.csv'

# The ID and range of a GNPS Enrolments spreadsheet
ENROLMENTS_SPREADSHEET_ID = '1TWHLO0tfbiJJapuOfoTO2UWPGSbOHtFrMyS_nz5MQsA'
ENROLMENTS_RANGE_NAME = 'Form Responses!A1:AZ'

# The function to convert data from spreadsheets to data frame
def gsheet2df(gsheet):
    header = gsheet.get('values', [])[0]   # Assumes first line is header!
    values = gsheet.get('values', [])[1:]  # Everything else is data.
    if not values:
        print('No data found.')
    else:
        all_data = []
        for col_id, col_name in enumerate(header):
            column_data = []
            for row in values:
                column_data.append(row[col_id])
            ds = pd.Series(data=column_data, name=col_name)
            all_data.append(ds)
        df = pd.concat(all_data, axis=1)
        return df

def main():
    # spreadsheet_service = build('sheets', 'v4', credentials=creds)
    directory_service = services.directoryService
    spreadsheet_service = services.spreadsheetService
    # Call the Google Sheets API
    enrolments_result = spreadsheet_service.spreadsheets().values().get(spreadsheetId=ENROLMENTS_SPREADSHEET_ID,range=ENROLMENTS_RANGE_NAME).execute()
    # Call the Google Admin API
    page_token = None
    total_users = []
    while True:
        try:
            users_result = directory_service.users().list(customer='my_customer', maxResults=500, pageToken=page_token).execute()
            users = users_result.get('users', [])
            total_users = [*total_users, *users] 
            page_token = users_result.get('nextPageToken')
            if not page_token:
                break 
        except errors.HttpError:
            break
        
    df = gsheet2df(enrolments_result)
    enrolments = df.values.tolist()
    output_header = ["submission_date", "class_format", "existing_student","student_id","name_on_system", "name_on_registration", "father_name","father_mobile", "mother_name","mother_mobile"]
    with open(_outputFolder/_outputFileName, 'w', newline='') as file:
        writer = csv.writer(file)
        writer.writerow(output_header)
        for enrolment in enrolments:
            if (enrolment[3] == 'Yes'):
                res = next((sub for sub in total_users if (sub['primaryEmail'] == enrolment[4] + '@gnps.nsw.edu.au')), None)
                # if ID is found but the first name is a mismatch 
                if res != None  and str(enrolment[5]).upper().rsplit(' ')[0] in str(res['name']['givenName']).upper():
                    print('%s,%s,%s,%s' % (enrolment[4], enrolment[5], str(res['name']['fullName']), str(enrolment[5] + ' ' + enrolment[6] + ' ' + enrolment[7])))
                    writer.writerow([enrolment[0], enrolment[2], enrolment[3], enrolment[4], str(res['name']['fullName']).upper().strip(), str(enrolment[5].upper().strip() + ' ' + enrolment[6].upper().strip() + ' ' + enrolment[7].upper().strip()).upper(),str(enrolment[14] + ' ' + enrolment[15]), enrolment[16] if str(enrolment[16]).startswith('0') else str('0' + enrolment[16]),str(enrolment[18] + ' ' + enrolment[19]), enrolment[20] if str(enrolment[20]).startswith('0') else str('0' + enrolment[20])  ])
                # if no ID match is found
                else:
                    print('%s,%s,%s' % (enrolment[4], enrolment[5], 'No Match'))
                    #writer.writerow([enrolment[0], enrolment[2], enrolment[3], enrolment[4], str(enrolment[5] + ' ' + enrolment[6] + ' ' + enrolment[7]).upper() + ' (POTENTIAL ISSUE WITH REGISTRATION)', str(enrolment[5] + ' ' + enrolment[6] + ' ' + enrolment[7]),str(enrolment[14] + ' ' + enrolment[15]), enrolment[16] if str(enrolment[16]).startswith('0') else str('0' + enrolment[16]),str(enrolment[18] + ' ' + enrolment[19]), enrolment[20] if str(enrolment[20]).startswith('0') else str('0' + enrolment[20])  ])
                    writer.writerow([enrolment[0], enrolment[2], enrolment[3], enrolment[4], str(res['name']['fullName']).upper().strip() + ' (POTENTIAL ISSUE WITH REGISTRATION)', str(enrolment[5] + ' ' + enrolment[6] + ' ' + enrolment[7]),str(enrolment[14] + ' ' + enrolment[15]), enrolment[16] if str(enrolment[16]).startswith('0') else str('0' + enrolment[16]),str(enrolment[18] + ' ' + enrolment[19]), enrolment[20] if str(enrolment[20]).startswith('0') else str('0' + enrolment[20])  ])

    print('Total records checked: ', len(enrolments))
    print('Please check the generated out file: '+ _outputFileName)
if __name__ == '__main__':
    main()