from __future__ import print_function
import sys
# appending a path
import lib.auth as auth
import lib.services as services
from pathlib import Path

# import specific modules
import pandas as pd
import random
import string
import re
import csv
from datetime import date
#output folder
_outputFolder = Path("output")
_outputFileName = 'insert_student_log.csv'

# initiate Google Service APIs
spreadsheet_service = services.spreadsheetService
directory_service = services.directoryService
# The ID and range of a GNPS Enrolments spreadsheet
STUDENTS_SPREADSHEET_ID = '1W7pY-5-gr3rZTXnhz4-Ftjblzv44eOS6o8mLRO6tzLs'
STUDENTS_RANGE_NAME = 'ONLINE'

def connect():
    global classroom_service, spreadsheet_service
    spreadsheet_service = services.spreadsheetService.spreadsheets()
    classroom_service = services.classroomService
#
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
#
def addStudent():
    students_result = spreadsheet_service.values().get(
        spreadsheetId=STUDENTS_SPREADSHEET_ID, range=STUDENTS_RANGE_NAME).execute()
    df = gsheet2df(students_result)
    gnps_students = df.values.tolist()
    #
    log_output_header = ["Student ID", "Name", "Result"]
    
    # with open(_outputFolder/_outputFileName, 'w', newline='') as file:
    #     writer = csv.writer(file)
    #     writer.writerow(log_output_header)
    for gnps_student in gnps_students:
        if (gnps_student[3] == 'NR44'):
            if (gnps_student[8] != ''
            and gnps_student[9] != ''
            and gnps_student[10] != ''
            and gnps_student[11] != ''):
                student = {
                    'userId': gnps_student[4] + '@gnps.nsw.edu.au'
                }
                try:
                    student = classroom_service.courses().students().create(
                        courseId=gnps_student[10], enrollmentCode=gnps_student[11], body=student).execute()
                    print(student.get('profile').get('name').get(
                        'fullName') + ':' + gnps_student[4])
                except Exception as e:
                    print(e)
                    pass

def main():
    connect()
    addStudent()

if __name__ == '__main__':
    main()
