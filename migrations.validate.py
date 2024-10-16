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
_outputFileName = 'student_migration_registration_check.csv'

# The ID and range of a GNPS Enrolments spreadsheet
MIGRATIONS_SPREADSHEET_ID = '1vWCf2n7h71fQaK0qHPCLLDIcIox3PzWwjXa6WpTJzOo'
MIGRATIONS_RANGE_NAME = 'Migrations!A1:AQ'

# The ID and range of input Enrolments spreadsheet
ENROLMENTS_SPREADSHEET_ID = '1TWHLO0tfbiJJapuOfoTO2UWPGSbOHtFrMyS_nz5MQsA'
ENROLMENTS_RANGE_NAME = 'Form Responses!A2:AZ'
#
MIGRATIONS_CHECKED_RANGE_NAME = 'Migrations!A1:AQ'
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
    sheet =  spreadsheet_service.spreadsheets()
    migrations_result = sheet.values().get(spreadsheetId=MIGRATIONS_SPREADSHEET_ID,
                                range=MIGRATIONS_RANGE_NAME).execute()
    enrolments_result = sheet.values().get(spreadsheetId=ENROLMENTS_SPREADSHEET_ID,
                            range=ENROLMENTS_RANGE_NAME).execute()
    migrations_df = gsheet2df(migrations_result)
    enrolments_df = gsheet2df(enrolments_result)
    # df.to_csv (r'export_dataframe.csv', index = False, header=True)
    migrations = migrations_df.values.tolist()
    enrolments = enrolments_df.values.tolist()
    my_enrolments = []
    for enrolment in enrolments:
        my_enrolments.append(enrolment[4]) 
    #
    output_header = ["student_id","first_name","last_name","enrolment_name", "status","existing_format", "new_format"]
    with open(_outputFolder/_outputFileName, 'w', newline='') as file:
        writer = csv.writer(file)
        writer.writerow(output_header)
        student_new_format = ''
        student_existing_format = ''
        for migration in migrations:
            if migration[0] in my_enrolments:
                for enrolment in enrolments:
                    if enrolment[4] == migration[0]:
                        enrolment_name = enrolment[5] + ' ' + enrolment[7]
                        student_new_format = enrolment[2]
                        student_existing_format = migration[3]
                print(migration[0],migration[1], migration[2],'Found Registration')
                writer.writerow([migration[0],migration[1], migration[2],enrolment_name,'Yes',student_existing_format,student_new_format])
            else:
                print(migration[0],migration[1], migration[2],'No Registration Found')
                writer.writerow([migration[0],migration[1], migration[2],'NO MATCH','No',student_existing_format,'NOT DETERMINED'])

if __name__ == '__main__':
    main()