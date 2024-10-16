import gspread
from oauth2client.service_account import ServiceAccountCredentials
from googleapiclient.discovery import build
import pandas as pd

# set up credentials to access Google Sheets and Google Classroom APIs
scope = ['https://www.googleapis.com/auth/classroom.courses.readonly', 
         'https://www.googleapis.com/auth/classroom.rosters.readonly', 
         'https://www.googleapis.com/auth/spreadsheets']

creds = ServiceAccountCredentials.from_json_keyfile_name('common/authentication/credentials.json', scope)
gc = gspread.authorize(creds)
service = build('classroom', 'v1', credentials=creds)

# get a list of all courses
results = service.courses().list().execute()
courses = results.get('courses', [])

# set up a Google Sheet to store the student data
worksheet_name = 'Classroom Students'
sh = gc.create(worksheet_name)

# iterate through each course and create a separate sheet for each
for course in courses:
    sheet_name = course.get('name') + ' Students'
    worksheet = sh.add_worksheet(title=sheet_name, rows=100, cols=20)
    worksheet.append_row(['Name', 'Email'])

    # get a list of all students in the course
    students = service.courses().students().list(courseId=course.get('id')).execute().get('students', [])
    
    #set up data structure to contain the students data
    sheet_data = []
    
    # For each student in the course, add their name and email to the data structure
    for student in students:
        profile = student.get('profile', {})
        name = ' '.join([profile.get('givenName', ''), profile.get('familyName', '')])
        email = profile.get('emailAddress', '')
        sheet_data.append([name, email])
    
    # Write the data to the newly created sheet in the Google Sheet
    for row in sheet_data:
        worksheet.append_row(row)

print('Classroom student data has been successfully written to a Google Sheet')