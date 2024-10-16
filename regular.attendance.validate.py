from __future__ import print_function
import sys
import random
# appending a path
import lib.auth as auth
import lib.services as services
from pathlib import Path
# import specific modules
import pandas as pd
import csv
#output folder
_outputFolder = Path("output")
_outputFileName = 'REGULAR_ATTENDANCE_T4.csv'
# The ID and range of a GNPS PCG2023 spreadsheet
TERM_STUDENTS_SPREADSHEET_ID = '1HkPDpPGAFo-YLzYZ5vT7Y9LC8v7TpSX2qoF-T0haNVg'
#
TERM_STUDENTS_RANGE_NAME = 'ALL!A1:AZ'

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

def generate_string(num_P):
    if num_P > 2:
        num_P = 2  # Cap the number of P's at 8
    num_A = 2 - num_P  # Determine the number of A's based on the number of P's
    # Create a list with the specified number of P's and A's
    char_list = ['P'] * num_P + ['A'] * num_A
    # Shuffle the list to randomize the order of characters
    random.shuffle(char_list)
    # Join the characters in the list to form a string, with each character separated by a comma
    result_string = ','.join(char_list)
    return result_string

def main():
    # spreadsheet_service = build('sheets', 'v4', credentials=creds)
    spreadsheet_service = services.spreadsheetService
    # Call the Google Sheets API
    sheet =  spreadsheet_service.spreadsheets()
    TERM_result = sheet.values().get(spreadsheetId=TERM_STUDENTS_SPREADSHEET_ID,
                                range=TERM_STUDENTS_RANGE_NAME).execute()
    TERM_df = gsheet2df(TERM_result)
    # -----------------------------------
    TERM_RECORDS = TERM_df.values.tolist()
    output_header = ["student_id","first_name","last_name","date_of_birth","gender","attendance","allergies","teacher","wk1","wk2","wk3","wk4","wk5","wk6","wk7","wk8","wk9"]
    with open(_outputFolder/_outputFileName, 'w', newline='') as file:
        writer = csv.writer(file)
        writer.writerow(output_header)
        #
        for termRecord in TERM_RECORDS:
            attendance = random.randint(1, 2)
            #print(termRecord[4],termRecord[18], termRecord[20], match,attendance,generate_string(int(attendance)))
            my_row = [ termRecord[4], #student_id
                        termRecord[18], # first name
                        termRecord[20], # last name
                        termRecord[21], #dob
                        termRecord[22], # gender
                        attendance, # attendance
                        termRecord[37],
                        termRecord[12]
                     ]
            attendance_row = generate_string(int(attendance)).split(",")
            print(my_row + attendance_row)
            writer.writerow(my_row + attendance_row)
if __name__ == '__main__':
    main()