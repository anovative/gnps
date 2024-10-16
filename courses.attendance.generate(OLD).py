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
# ATTENDANCE_SPREADSHEET_ID = '1TXqHnq1mD1M1DEWMWsUcnnbYHcUemKv5FMzeadGgMz8'
ATTENDANCE_SPREADSHEET_ID = '1E1A7q7Q5cic2zHME1sAAtXIjsXBRGlrrLMYaNeMqgyY'
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
            if course.get('section') != None : 
                print('Course ID:%s Course Name:%s Teacher(s):%s' %(course.get('id'),course.get('name'), course.get('section')))
                if (course.get('name') + '-' + course.get('section')) != '':
                    request = spreadsheet_service.spreadsheets().get(spreadsheetId=ATTENDANCE_SPREADSHEET_ID, ranges=course.get('name') + '-' + course.get('section'), includeGridData=False)
                    response = request.execute()
                    print(response['sheets'][0]['properties'])
                    if (response['sheets'][0]['properties']['title'] == course.get('name') + '-' + course.get('section')):
                        request_body_ext = {
                                    "requests": [
                                            {
                                                "deleteSheet": {
                                                    "sheetId": response['sheets'][0]['properties']['sheetId']
                                                }
                                            }
                                        ]
                                    }
                        spreadsheet_service.spreadsheets().batchUpdate(
                            spreadsheetId=ATTENDANCE_SPREADSHEET_ID,
                            body=request_body_ext
                        ).execute()
                        request_body = {
                                    'requests': [
                                        {
                                            'addSheet': {
                                                'properties': {
                                                    'title': course.get('name') + '-' + course.get('section'),
                                                }
                                            },
                                        },
                                                ]
                        }
                        spreadsheet_service.spreadsheets().batchUpdate(
                            spreadsheetId=ATTENDANCE_SPREADSHEET_ID,
                            body=request_body
                        ).execute()
                    else:
                        request_body = {
                                    'requests': [
                                        {
                                            'addSheet': {
                                                'properties': {
                                                    'title': course.get('name') + '-' + course.get('section'),
                                                }
                                            },
                                        },
                                                ]
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
                            try:
                                print('Student ID: %s  Student Name: %s' % (student.get('profile')['emailAddress'].split('@')[0],student.get('profile')['name']['fullName']))
                                student_records.append([student.get('profile')['emailAddress'].split('@')[0],student.get('profile')['name']['fullName']])
                            except Exception as e:
                                print(e)
                                pass
                    else:
                        print('No Students found in this course')
                    #print(response)
                    #print(response['sheets'][0]['properties'])
                    #print(response['sheets'][0]['properties']['title'] + ':' + str(response['sheets'][0]['properties']['sheetId']))
                    student_rec = {
                        "majorDimension": "ROWS",
                        "values": student_records
                    }
                    spreadsheet_service.spreadsheets().values().update(
                    spreadsheetId=ATTENDANCE_SPREADSHEET_ID,
                    range=course.get('name') + '-' + course.get('section'),
                    valueInputOption='USER_ENTERED',
                    body=student_rec
                    ).execute()
        #iterate += 1
        
if __name__ == '__main__':
    main()