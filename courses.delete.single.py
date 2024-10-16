from __future__ import print_function
from ast import IsNot
import lib.auth as auth
import lib.services as services
from copy import Error

def connect():
    global classroom_service, spreadsheet_service
    classroom_service = services.classroomService
#
def deleteCourses()->None:
    try:
        classrooms = classroom_service.courses()
        classroom = classrooms.delete(id=188517961443).execute()
    except Error as error:
        print('Error Deleting')
    
def main():
    connect()
    deleteCourses()

if __name__ == '__main__':
    main()