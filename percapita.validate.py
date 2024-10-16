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
_outputFileName = 'F2F-STUDENTS(NEW)_student_pcg2023record_registration_check.csv'

# The ID and range of a GNPS PCG2023 spreadsheet
PCG2022_SPREADSHEET_ID = '1-ZJUUd7wLVOo4U0xoHMyuEJ7TnXVFev6nfRkogxpniw'
PCG2022_RANGE_NAME = 'RECORDS!A1:F'

# The ID and range of input PCG2023 spreadsheet
PCG2023_SPREADSHEET_ID = '1UUSipEDlCyx5Ay1zmFiJXv4smHOb5PpzEx-u2wI72-E'

# PCG2023_RANGE_NAME = 'ONL-STUDENTS(NEW)!A1:AZ'
# PCG2023_RANGE_NAME = 'ONL-STUDENTS(EXISTING)!A1:AZ'
# PCG2023_RANGE_NAME = 'F2F-STUDENTS(NEW)!A1:AZ'
PCG2023_RANGE_NAME = 'F2F-STUDENTS(NEW)!A1:AZ'
#
PCG2022_CHECKED_RANGE_NAME = 'PCG2022!A1:AQ'
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
    spreadsheet_service = services.spreadsheetService
    # Call the Google Sheets API
    sheet =  spreadsheet_service.spreadsheets()
    PCG2022_result = sheet.values().get(spreadsheetId=PCG2022_SPREADSHEET_ID,
                                range=PCG2022_RANGE_NAME).execute()
    PCG2023_result = sheet.values().get(spreadsheetId=PCG2023_SPREADSHEET_ID,
                            range=PCG2023_RANGE_NAME).execute()
    PCG2022_df = gsheet2df(PCG2022_result)
    PCG2023_df = gsheet2df(PCG2023_result)
    # -----------------------------------
    PCG2022 = PCG2022_df.values.tolist()
    PCG2023 = PCG2023_df.values.tolist()
    output_header = ["first_name","last_name","date_of_birth","gender","attendance","carer_mobile","mainstream_school","mainstream_school_year","match"]
    with open(_outputFolder/_outputFileName, 'w', newline='') as file:
        writer = csv.writer(file)
        writer.writerow(output_header)
        for pcg2023record in PCG2023:
            match = 'No'
            for pcg2022record in PCG2022:
                if (pcg2023record[18] == pcg2022record[0]
                   and pcg2023record[20] == pcg2022record[1]
                   and pcg2023record[21] == pcg2022record[2]
                   and pcg2023record[22] == pcg2022record[3]
                   ):
                        match = 'Yes'
                elif (pcg2023record[18] == pcg2022record[0]
                    and pcg2023record[20] == pcg2022record[1]
                    and pcg2023record[22] == pcg2022record[3]
                    ):
                        match = 'May'
            print(pcg2023record[18], pcg2023record[20], match)
            writer.writerow(
                            [
                             pcg2023record[18], # first name
                             pcg2023record[20], # last name
                             pcg2023record[21].replace("-", "/" ), # dob
                             pcg2023record[22], # gender
                             random.randint(6,8), # attendance
                             pcg2023record[27],
                             pcg2023record[23],
                             pcg2023record[24],
                             match # match/no match
                             ]
                            )
if __name__ == '__main__':
    main()