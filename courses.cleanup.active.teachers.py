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
        print('Course:', activeCourse)
        teachers = []
        page_token = None
        while True:
            response = classroom_service.courses().teachers().list(courseId=activeCourse,pageToken=page_token,
                                            pageSize=100).execute()
            teachers.extend(response.get('teachers', []))
            page_token = response.get('nextPageToken', None)
            if not page_token:
                break
        if teachers:
            for teacher in teachers:
                try:
                    if teacher.get('userId') != '103698760749353484985':
                        print('Removing Teacher with ID: %s Name: %s' % (teacher.get('userId'),teacher.get('profile')['name']['fullName']))
                        classroom_service.courses().teachers().delete(courseId=activeCourse, userId=teacher.get('userId')).execute()
                    # print(teacher)
                except Exception as e:
                    print(e)
                    pass
        else:
            print('No Teachers found in this course')
if __name__ == '__main__':
    main()