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
STUDENTS_SPREADSHEET_ID = '1zToCwZB1vrpH4oOINY4RPa_yU_FeC_NC73ThKw0glOA'
STUDENTS_RANGE_NAME = 'Main'

_outputFolder = Path("output/students/")

def connect():
    global classroom_service, spreadsheet_service, directory_service
    spreadsheet_service = services.spreadsheetService.spreadsheets()
    classroom_service = services.classroomService
    directory_service = services.directoryService

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


def listCourses(studentId):
    courses = []
    page_token = None
    while True:
        response = classroom_service.courses().list(studentId=studentId).execute()
        courses.extend(response.get('courses', []))
        page_token = response.get('nextPageToken', None)
        if not page_token:
            break
    return courses
#


def findStudent():
    students_result = spreadsheet_service.values().get(
        spreadsheetId=STUDENTS_SPREADSHEET_ID, range=STUDENTS_RANGE_NAME).execute()
    df = gsheet2df(students_result)
    #
    page_token = None
    total_users = []
    while True:
        try:
            users_result = directory_service.users().list(
                customer='my_customer', maxResults=500, pageToken=page_token).execute()
            users = users_result.get('users', [])
            total_users = [*total_users, *users]
            page_token = users_result.get('nextPageToken')
            if not page_token:
                break
        except errors.HttpError:
            break
    #
    gnps_students = df.values.tolist()
    log_output_header = ["StudentID", "StudentName",
                         "CourseID", "CourseName", "CourseSection", "CourseState"]
    with open(_outputFolder/'remove_student_log.csv', 'w', newline='') as file:
        writer = csv.writer(file)
        writer.writerow(log_output_header)
        for gnps_student in gnps_students:
            student = {
                'userId': gnps_student[0] + '@gnps.nsw.edu.au'
            }
            try:
                res = next((sub for sub in total_users if (
                    sub['primaryEmail'] == gnps_student[0] + '@gnps.nsw.edu.au')), None)
                print(gnps_student[0] + '@gnps.nsw.edu.au : ' +
                      str(res['name']['fullName']))
                courses = listCourses(gnps_student[0] + '@gnps.nsw.edu.au')
                # print(courses)
                if not courses:
                    print('No courses found.')
                else:
                    for course in courses:
                        #print('%s,%s,%s' % (course.get('name'),course.get('section'),course.get('id')))
                        student = classroom_service.courses().students().delete(
                            courseId=course.get('id'), userId=gnps_student[0] + '@gnps.nsw.edu.au').execute()
                        writer.writerow([gnps_student[0], str(res['name']['fullName']), course.get(
                            'id'), course.get('name'), course.get('section'), course.get('courseState')])
            except Exception as e:
                # print("Error: " + gnps_student[0])
                print(e)
                writer.writerow([gnps_student[0], 'Error', 'Error', 'Error'])
                pass


def main():
    connect()
    findStudent()
    # listCourses()


if __name__ == '__main__':
    main()
