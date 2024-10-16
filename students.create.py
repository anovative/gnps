from __future__ import print_function
from ast import IsNot
import lib.auth as auth
import lib.services as services
from pathlib import Path
from copy import Error
# 
import pandas as pd
import csv
import time
import os.path
import random
import string
#
#EXECUTION_MODE = "ONLINE"
EXECUTION_MODE = "F2F"

#output folder
_outputFolder = Path("output/students/"+ EXECUTION_MODE)
_newUserOutputFileName = EXECUTION_MODE + '_new_users.csv'
_existingUserOutputFileName = EXECUTION_MODE + '_existing_users.csv'
_groupMembershipOutputFileName = EXECUTION_MODE + '_group_membership.csv'
_combinedOutputFileName = EXECUTION_MODE + '_combined_update.csv'

# The ID and range of a GNPS Enrolments spreadsheet
ENROLMENTS_SPREADSHEET_ID = '1W7pY-5-gr3rZTXnhz4-Ftjblzv44eOS6o8mLRO6tzLs' 
ENROLMENTS_RANGE_NAME = EXECUTION_MODE  + '!A1:BB'

def connect():
    global classroom_service, spreadsheet_service
    spreadsheet_service = services.spreadsheetService.spreadsheets()
    classroom_service = services.classroomService
    
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
    #sheet = spreadsheet_service.spreadsheets()
    enrolments_result = spreadsheet_service.values().get(spreadsheetId=ENROLMENTS_SPREADSHEET_ID,
                                range=ENROLMENTS_RANGE_NAME).execute()
    # Call the Google Admin API
    df = gsheet2df(enrolments_result)
    # df.to_csv (r'export_dataframe.csv', index = False, header=True)
    enrolments = df.values.tolist()
    new_output_header = ["First Name [Required]","Last Name [Required]","Email Address [Required]","Password [Required]","Org Unit Path [Required]","Change Password at Next Sign-In","Home Phone","Mobile Phone"]
    existing_output_header = ["First Name [Required]","Last Name [Required]","Email Address [Required]","Org Unit Path [Required]","Home Phone","Mobile Phone"]
    group_output_header = ["Group Email [Required]","Member Email","Member Type","Member Role"]
    #
    #with open(_outputFolder/'students/' + EXECUTION_MODE + '/'+ EXECUTION_MODE + '_new_users.csv', 'w', newline='') as file:
    with open(_outputFolder/_newUserOutputFileName, 'w', newline='') as file:
        writer = csv.writer(file)
        writer.writerow(new_output_header)
        for enrolment in enrolments:
            # letters_and_digits = string.ascii_letters + string.digits
            # result_str = ''.join((random.choice(letters_and_digits) for i in range(8)))
            if (enrolment[3] == 'NR22'):
                if (enrolment[1] == 'No'):
                    writer.writerow([enrolment[18], enrolment[20], enrolment[51], enrolment[52], enrolment[53], 'FALSE', enrolment[27], enrolment[31]])
        #
    #with open(_outputFolder/'students/' + EXECUTION_MODE + '/' + EXECUTION_MODE + '_existing_users.csv', 'w', newline='') as file:
    with open(_outputFolder/_existingUserOutputFileName, 'w', newline='') as file:
        writer = csv.writer(file)
        writer.writerow(existing_output_header)
        for enrolment in enrolments:
            if (enrolment[3] == 'NR22'):
                if (enrolment[1] == 'Yes'):
                    writer.writerow([enrolment[18], enrolment[20], enrolment[51],enrolment[53],enrolment[27], enrolment[31]])
    #
    #with open(_outputFolder/'students/' + EXECUTION_MODE + '/'+ EXECUTION_MODE + '_group_membership.csv', 'w', newline='') as file:
    with open(_outputFolder/_groupMembershipOutputFileName, 'w', newline='') as file:
        writer = csv.writer(file)
        writer.writerow(group_output_header)
        for enrolment in enrolments:
            if (enrolment[3] == 'NR22'):
                writer.writerow([enrolment[17], enrolment[51], 'USER','MEMBER'])
    #
    #with open(_outputFolder/'students/' + EXECUTION_MODE + '/' + EXECUTION_MODE + '_existing_users.csv', 'w', newline='') as file:
    with open(_outputFolder/_combinedOutputFileName, 'w', newline='') as file:
        writer = csv.writer(file)
        writer.writerow(existing_output_header)
        for enrolment in enrolments:
            if (enrolment[3] == 'NR22'):
                if (enrolment[1] == 'Yes' or enrolment[1] == 'No'):
                    writer.writerow([enrolment[18], enrolment[20], enrolment[51],enrolment[53],enrolment[27], enrolment[31]])
        print('Total records found: ', len(enrolments))
if __name__ == '__main__':
    connect()
    main()