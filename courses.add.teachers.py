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

# The ID and range of a GNPS Enrolments spreadsheet
TEACHERS_SPREADSHEET_ID = '1pR5Eyi7fbxTVk35k2slntdML3moiCVp1NFEvmOY7b0o'
TEACHERS_RANGE_NAME = 'Courses(Created)-W!A1:AZ'
#
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
def addTeachers():
    teachers_result = spreadsheet_service.values().get(
        spreadsheetId=TEACHERS_SPREADSHEET_ID, range=TEACHERS_RANGE_NAME).execute()
    df = gsheet2df(teachers_result)
    gnps_teachers = df.values.tolist()
    #

    for gnps_teacher in gnps_teachers:
        teacher = {
            'userId': gnps_teacher[4]
        }
        try:
            classroom_service.courses().teachers().create(
                courseId=gnps_teacher[1], body=teacher).execute()
            #
            course = classroom_service.courses().get(id=gnps_teacher[1]).execute()
            course['section'] = gnps_teacher[3]
            course = classroom_service.courses().update(id=gnps_teacher[1], body=course).execute()
            #
            print('Added %s to: %s' % (gnps_teacher[3],gnps_teacher[1]))
            time.sleep(3)
        except Error as error:
            print('Error Occurred: ' + gnps_teacher[3])
def main():
    connect()
    addTeachers()
#
if __name__ == '__main__':
    main()
