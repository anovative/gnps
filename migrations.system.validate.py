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
_outputFileName = 'migrations.system_check_results.csv'

MIGRATIONS_SPREADSHEET_ID = '1vWCf2n7h71fQaK0qHPCLLDIcIox3PzWwjXa6WpTJzOo'
MIGRATIONS_RANGE_NAME = 'Migrations!A1:AQ'

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
    migrations_result = spreadsheet_service.spreadsheets().values().get(spreadsheetId=MIGRATIONS_SPREADSHEET_ID,range=MIGRATIONS_RANGE_NAME).execute()
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
        
    df = gsheet2df(migrations_result)
    migrations = df.values.tolist()
    output_header = ["student_id","name_on_system", "name_on_migration", "status"]
    with open(_outputFolder/_outputFileName, 'w', newline='') as file:
        writer = csv.writer(file)
        writer.writerow(output_header)
        for migration in migrations:
            res = next((sub for sub in total_users if (sub['primaryEmail'] == migration[0] + '@gnps.nsw.edu.au')), None)
            # if ID is found but the first name is a mismatch 
            if res != None  and str(migration[1]).upper().rsplit(' ')[0] in str(res['name']['givenName']).upper():
                print('%s,%s,%s,%s' % (migration[0], migration[1], str(res['name']['fullName']), str(migration[1] + ' ' + migration[2])))
                writer.writerow([migration[0], str(res['name']['fullName']).upper().strip(), str(migration[1].upper().strip() + ' ' + migration[2].upper().strip()).upper(),'MATCHED'])
            # if no ID match is found
            else:
                print('%s,%s,%s' % (migration[0], migration[1], 'No Match'))
                writer.writerow([migration[0], str(res['name']['fullName']).upper().strip(), str(migration[1].upper().strip() + ' ' + migration[2].upper().strip()).upper(),'POTENTIAL ISSUE WITH MIGRATION'])

    print('Total records checked: ', len(migrations))
    print('Please check the generated out file: '+ _outputFileName)
if __name__ == '__main__':
    main()