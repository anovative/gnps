from __future__ import print_function
from ast import IsNot
import lib.auth as auth
import lib.services as services
from pathlib import Path
# //
import pandas as pd
import csv
import time
import os.path
#
# The ID and range of a GNPS Enrolments spreadsheet
COURSES_SPREADSHEET_ID = '1pR5Eyi7fbxTVk35k2slntdML3moiCVp1NFEvmOY7b0o'
COURSES_RANGE_NAME = 'Courses(TBC)!A1:AZ'
CREATED_COURSES_RANGE_NAME = 'Courses(Created)!A1:AZ'
OUTPUT_HEADER = ['course_name','course_id','course_code','course_teacher_name','course_teacher_email']
ADMIN_TEACHER = "sukhvinder.flora@gnps.nsw.edu.au"
#
courseRecords = []
courseRecords.append(OUTPUT_HEADER)
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
def addCourses():
    courseResult = spreadsheet_service.values().get(
        spreadsheetId=COURSES_SPREADSHEET_ID, range=COURSES_RANGE_NAME).execute()
    df = gsheet2df(courseResult)
    gnps_classrooms = df.values.tolist()
    #
    for gnps_classroom in gnps_classrooms:
        course = {
            'name': gnps_classroom[0] + '(' + gnps_classroom[1] + ')' ,
            'description': gnps_classroom[2],
            'room': gnps_classroom[3],
            'ownerId': 'me',
            'courseState': 'ACTIVE'
        }
        try:
            course = classroom_service.courses().create(body=course).execute()
            teachers = classroom_service.courses().teachers()
            print('Course created: %s %s %s' % (course.get('name'), course.get('id'), course.get('enrollmentCode')))
            time.sleep(3)
            teacher = {
                'userId': ADMIN_TEACHER
            }
            teacher = teachers.create(
                     courseId=course.get('id'), body=teacher).execute()
            courseRecords.append([course.get('name'),course.get('id'), course.get('enrollmentCode')])
        except Error as error:
            print('Error Creating: ' + gnps_classroom[0])
    classroom_resource = {
        "majorDimension": "ROWS",
        "values": courseRecords
    }
    spreadsheet_service.values().append(
    spreadsheetId=COURSES_SPREADSHEET_ID,
    range=CREATED_COURSES_RANGE_NAME,
    valueInputOption='USER_ENTERED',
    insertDataOption = 'OVERWRITE',
    body=classroom_resource
    ).execute()
#
def main():
    connect()
    addCourses()
#
if __name__ == '__main__':
    main()
