from __future__ import print_function
from ast import IsNot
import lib.auth as auth
import lib.services as services
from pathlib import Path
#
def connect():
    global classroom_service, spreadsheet_service
    classroom_service = services.classroomService
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
    for activeCourse in getActiveCourses():
        try:
            print('Removing Course with ID: %s' % (activeCourse))
            classroom_service.courses().delete(id=activeCourse).execute()
        except Exception as e:
            print(e)
            pass
if __name__ == '__main__':
    main()