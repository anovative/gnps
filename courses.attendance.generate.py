from __future__ import print_function
from ast import IsNot
import lib.auth as auth
import lib.services as services
from pathlib import Path
#
def connect():
    global classroom_service, spreadsheet_service
    classroom_service = services.classroomService
    spreadsheet_service = services.spreadsheetService
    
# The ID and range of input Enrolments spreadsheet
ATTENDANCE_SPREADSHEET_ID = '1ihM6zgqv9pZqlmKCXDfSINdde_wey7g9OSR2tJ6a7Io'
OUTPUT_HEADER = ['ID','Name']
#
def getActiveCourses()->None:
    activeCourses = []
    courses = []
    page_token = None
    while True:
        response = classroom_service.courses().list(pageToken=page_token,
                                          pageSize=100).execute()
        courses.extend(response.get('courses', []))
        page_token = response.get('nextPageToken', None)
        if not page_token:
            break
    for course in courses:
        if (course.get('courseState') != 'ARCHIVED'
            and course.get('courseState') != 'PROVISIONED'
            and course.get('courseState') != 'DECLINED'
            and course.get('courseState') != 'SUSPENDED'
            and course.get('id') != '91527479820' # ignore the internal teacher training course
            ):
            activeCourses.append(course.get('id'))
    return activeCourses


def main():
    connect()
    iterate = 1
    for activeCourse in getActiveCourses():
        #if iterate <= 3: # control the loop
        course = classroom_service.courses().get(id=activeCourse).execute()
        student_records = []
        student_records.append(['Class','Teacher','Room'])
        student_records.append([course.get('name'), course.get('section'), course.get('room')])
        student_records.append([''])
        student_records.append(['Students'])
        student_records.append(OUTPUT_HEADER)
        #course.get('section') 
        print('Course ID:%s Course Name:%s Teacher(s):%s' %(course.get('id'),course.get('name'), course.get('section')))
        # request_body_ext = {
        #             "requests": [
        #                 {
        #                 "deleteSheet": {
        #                     "sheetId": course.get('name') + '-' + course.get('section')
        #                 }
        #                 }
        #                 ]
        #             }
        # spreadsheet_service.spreadsheets().batchUpdate(
        #     spreadsheetId=ATTENDANCE_SPREADSHEET_ID,
        #     body=request_body_ext
        # ).execute()
        #
        request_body = {
                    'requests': [{
                        'addSheet': {
                            'properties': {
                                'title': course.get('name') + '-' + course.get('section') if course.get('section') != None else ' ',
                                #range=course.get('name') + '-' + course.get('section') if course.get('section') != None else ' ',
                                'tabColor': {
                                    'red': 0,
                                    'green': 255,
                                    'blue': 0
                                }
                            }
                        }
                    }]
        }
        spreadsheet_service.spreadsheets().batchUpdate(
            spreadsheetId=ATTENDANCE_SPREADSHEET_ID,
            body=request_body
        ).execute()
        students = []
        page_token = None
        while True:
            response = classroom_service.courses().students().list(courseId=activeCourse,pageToken=page_token,
                                            pageSize=100).execute()
            students.extend(response.get('students', []))
            page_token = response.get('nextPageToken', None)
            if not page_token:
                break
        if students:
            for student in students:
                #print(student)
                try:
                    print('Student ID: %s  Student Name: %s' % (student.get('profile')['emailAddress'].split('@')[0],student.get('profile')['name']['fullName']))
                    student_records.append([student.get('profile')['emailAddress'].split('@')[0],student.get('profile')['name']['fullName']])
                except Exception as e:
                    print(e)
                    pass
        else:
            print('No Students found in this course')
        student_rec = {
            "majorDimension": "ROWS",
            "values": student_records
        }
        spreadsheet_service.spreadsheets().values().update(
        spreadsheetId=ATTENDANCE_SPREADSHEET_ID,
        range=course.get('name') + '-' + course.get('section') if course.get('section') != None else '-',
        valueInputOption='USER_ENTERED',
        body=student_rec
        ).execute()
        #iterate += 1
        
if __name__ == '__main__':
    main()