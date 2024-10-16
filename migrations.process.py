from __future__ import print_function
import sys
# appending a path
import lib.auth as auth
import lib.services as services
from pathlib import Path
# import specific modules
import pandas as pd
import csv
import random
import string
#output folder
_outputFolder = Path("output")
_outputUsersFileName = 'migrating_users.csv'
_outputGroupsFileName = 'migrating_groups.csv'

#EXECUTION_MODE = "ONLINE"
EXECUTION_MODE = "F2F"

# The ID and range of a GNPS Enrolments spreadsheet
MIGRATIONS_SPREADSHEET_ID = '1j3N0xCePxtxZ1skwLNSciApeh0Y5byP6Q15lRbBkFS0'
MIGRATIONS_RANGE_NAME = 'Migrations!A1:AE'

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
    spreadsheet_service = services.spreadsheetService
    sheet = spreadsheet_service.spreadsheets()
    migrations_result = sheet.values().get(spreadsheetId=MIGRATIONS_SPREADSHEET_ID,
                                range=MIGRATIONS_RANGE_NAME).execute()
    # Call the Google Admin API
    df = gsheet2df(migrations_result)
    migrations = df.values.tolist()
    existing_output_header = ["First Name [Required]","Last Name [Required]","Email Address [Required]","Org Unit Path [Required]"]
    group_output_header = ["Group Email [Required]","Member Email","Member Type","Member Role"]
   #
    with open(_outputFolder/_outputUsersFileName, 'w', newline='') as file:
        writer = csv.writer(file)
        writer.writerow(existing_output_header)
        for migration in migrations:
            writer.writerow([migration[1].capitalize(), migration[2].capitalize(), migration[0]+'@gnps.nsw.edu.au','/Students/YEAR-2023/' + migration[3] + '/' + migration[5]])
    #
    with open(_outputFolder/_outputGroupsFileName, 'w', newline='') as file:
        writer = csv.writer(file)
        writer.writerow(group_output_header)
        for migration in migrations:
            writer.writerow([migration[4]+'_students@gnps.nsw.edu.au', migration[0]+'@gnps.nsw.edu.au', 'USER','MEMBER'])
    #
    print('Total records found: ', len(migrations))
if __name__ == '__main__':
    main()