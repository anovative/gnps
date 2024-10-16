from __future__ import print_function
from ast import IsNot
import lib.auth as auth
import lib.services as services
from pathlib import Path
# import specific modules
import pandas as pd
import csv
#output folder
_outputFolder = Path("output")
_outputFileName = 'classroom.list.active.csv'

def connect():
    global classroom_service, spreadsheet_service
    classroom_service = services.classroomService
#
def listCourses()->None:
    courses = []
    page_token = None
    while True:
        response = classroom_service.courses().list(pageToken=page_token,
                                          pageSize=100).execute()
        courses.extend(response.get('courses', []))
        page_token = response.get('nextPageToken', None)
        if not page_token:
            break

    if not courses:
        print('No courses found.')
    else:
        print('Courses:')
        output_header = ["course_name", "section", "enrolment_code","course_id","course_state"]
        with open(_outputFolder/_outputFileName, 'w', newline='') as file:
            writer = csv.writer(file)
            writer.writerow(output_header)
            for course in courses:
                if (course.get('courseState') != 'ACTIVE'):
                    print('%s,%s,%s,%s,%s' % (course.get('name'),course.get('section'),course.get('enrollmentCode'),course.get('id'),course.get('courseState')))
                    writer.writerow([course.get('name'),course.get('section'),course.get('enrollmentCode'),course.get('id'),course.get('courseState')])

def main():
    connect()
    listCourses()

if __name__ == '__main__':
    main()