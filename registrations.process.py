from __future__ import print_function
import sys
# appending a path
import lib.auth as auth
import lib.services as services
from pathlib import Path

# import specific modules
import random
import string
import re
from datetime import date

# initiate Google Service APIs
spreadsheet_service = services.spreadsheetService
directory_service = services.directoryService
# CONSTANTS
# The ID and range of input Enrolments spreadsheet
RAW_ENROLMENTS_SPREADSHEET_ID = '1TWHLO0tfbiJJapuOfoTO2UWPGSbOHtFrMyS_nz5MQsA'
RAW_ENROLMENTS_RANGE_NAME = 'Form Responses!A2:AZ'

# The ID and range of Output Enrolments spreadsheet
PROCESSED_ENROLMENTS_SPREADSHEET_ID = '1W7pY-5-gr3rZTXnhz4-Ftjblzv44eOS6o8mLRO6tzLs'
PROCESSED_ONLINE_ENROLMENTS_RANGE_NAME = 'ONLINE!A1:AZ'
PROCESSED_F2F_ENROLMENTS_RANGE_NAME = 'F2F!A1:AZ'
#
MIGRATIONS_SPREADSHEET_ID = '1vWCf2n7h71fQaK0qHPCLLDIcIox3PzWwjXa6WpTJzOo'
MIGRATIONS_RANGE_NAME = 'Migrations!A2:AQ'

#
ONLINE_CLASS_TIMINGS_SESSION1 = 'Saturday, 3:00pm-5:00pm'
ONLINE_CLASS_TIMINGS_SESSION2 = 'Saturday, 3:45pm-5:45pm'
#
F2F_CLASS_TIMINGS_SESSION1 = 'Saturday, 3:00pm-5:00pm'
F2F_CLASS_TIMINGS_SESSION2 = 'Saturday, 5:30pm-7:30pm'
#
GNPS_ID_SEED = 's-23-0003' ## IMPORTANT: This needs to be upadted with care with current value
OUTPUT_HEADER = ['Age-Group','Existing_Student','Format','Reason','Student_Id','Invoice_ID','Session','Course_Time','Room','Course_Name','Course_Id','Course_Code','Teacher(s)','Level','Level_Group','New_Level','Level_Type','New_Level_Group','Student_First_Name','Student_Middle_Name','Student_Surname','DOB','Gender','Regular_School','School_Year','Father_First_Name','Father_Last_Name','Father_mobile','Father_email','Mother_First_Name','Mother_Last_Name','Mother_mobile','Mother_email','Street','Suburb','State','Post_Code','Allergies','Medicare','Doctor','Doctor_Phone','EC1_First_Name','EC1_Last_Name','EC1_Phone','EC1_Relation','EC2_First_Name','EC2_Last_Name','EC2_Phone','EC2_Relation','Recipient','Email_Sent','User_Name','Password','Org_Path']
GNPS_DOMAIN = '@gnps.nsw.edu.au'
GNPS_ORG_PREFIX = '/Students/YEAR-2023/'
#
def calculate_age(user_birthday):
    #Get current date
    today = date.today()
    #
    split_dob = user_birthday.split("-")
    #Calculate the years difference
    year_diff = today.year - int(split_dob[2])
    #Check if birth month and birth day precede current month and current day
    precedes_flag = ((today.month, today.day) < (int(split_dob[1]), int(split_dob[0])))
    #Calculate the user's age
    age = year_diff - precedes_flag
    #Return the age value
    return age


def read_migrations():
    range_name = MIGRATIONS_RANGE_NAME
    spreadsheet_id = MIGRATIONS_SPREADSHEET_ID
    result = spreadsheet_service.spreadsheets().values().get(
    spreadsheetId=spreadsheet_id, range=range_name).execute()
    migrations = result.get('values', [])
    return migrations

#
def read_range():
    range_name = RAW_ENROLMENTS_RANGE_NAME
    spreadsheet_id = RAW_ENROLMENTS_SPREADSHEET_ID
    result = spreadsheet_service.spreadsheets().values().get(
    spreadsheetId=spreadsheet_id, range=range_name).execute()
    rows = result.get('values', [])
    return rows
#
page_token = None
total_users = []
while True:
    try:
        users_result = directory_service.users().list(customer='my_customer', maxResults=500, pageToken=page_token).execute()
        users = users_result.get('users', [])
        total_users = [*total_users, *users] 
        page_token = users_result.get('nextPageToken')
        if not page_token:
            break 
    except errors.HttpError:
        break
#
rows = read_range()
migrations = read_migrations()
online_records = []
f2f_records = []
#
online_records.append(OUTPUT_HEADER)
f2f_records.append(OUTPUT_HEADER)
#
level_group =''
GNPS_ID = GNPS_ID_SEED
for row in rows:
    if row[1] == 'PRCS':
        group_info = ''
        new_level = ''
        new_level_group = ''
        level_type = ''
        if row[3] != 'Yes': #if the studnets is NEW student
            # generate the students ID based on constant ID seed ( + increment)
            res = re.sub(r'[0-9]+$',lambda x: f"{str(int(x.group())+1).zfill(len(x.group()))}",GNPS_ID)
            GNPS_ID = str(res)
            #
            new_level = 'LEVEL-BASIC'
            new_level_group = 'level_basic_students' + GNPS_DOMAIN # if row[3]=='No' else group_info
            level_type = 'Standard'
        
        # generate random password for the new student
        letters_and_digits = string.ascii_letters + string.digits
        pwd_str = ''.join((random.choice(letters_and_digits) for i in range(8)))
        # for existing students, determine the group membership from the directory
        res = next((sub for sub in total_users if (sub['primaryEmail'] == row[4] + GNPS_DOMAIN)), None)
        #
        if row[3] != 'No' and res != None : #if the studnets is existing student
            groups = directory_service.groups().list(userKey=str(res['primaryEmail'])).execute().get('groups', [])
            for group in groups:
                group_info = group_info + str(group['email'])
            #
            for migration in migrations:
                new_level = migration[5]
                if migration[0] == row[4]:
                    # new_level = migration[5]
                    
                    # if migration[5].find('(ADV)') != -1:
                    #     level_type = migration[5][-5:]
                    # else:
                    #     level_type = migration[5]
                    level_type = 'Advanced' if migration[5].find('(ADV)') != -1 else 'Standard'
                    new_level_group = migration[6] + '_students' + GNPS_DOMAIN
                    break;
                else:
                    new_level = str(res['orgUnitPath']).rsplit("/")[len(str(res['orgUnitPath']).rsplit("/"))-1] if res != None else ''
                    new_level_group = level_group + GNPS_DOMAIN if row[3]=='No' else group_info
                    level_type = 'Standard'
        #
        match row[3]:
            case 'No':
                level_group = 'level_basic_students'
            case _:
                level_group = ''
        if row[2] == 'Online':
            if res != None :
                print(str(res['primaryEmail']))
            online_records.append([
                calculate_age(row[8]),
                row[3], # new or existing student
                row[2], # class format online/face-to-face
                row[48] if row[48] !='' else 'n/a', # reason for selecting online
                GNPS_ID if row[3]=='No' else row[4], # if existing student, list student ID
                row[47], # Invoice ID
                'Session-1' if row[3]=='No' else 'Session-2', # session
                ONLINE_CLASS_TIMINGS_SESSION1 if row[3]=='No' else ONLINE_CLASS_TIMINGS_SESSION2, #class time
                'n/a', # room
                '', # class code
                '', # class id
                '', # class name
                '', # Teachers
                'LEVEL-BASIC' if row[3]=='No' else str(res['orgUnitPath']).rsplit("/")[len(str(res['orgUnitPath']).rsplit("/"))-1] if res != None else '', # check and insert level from the system
                level_group + GNPS_DOMAIN if row[3]=='No' else group_info,
                new_level, # new level
                level_type, # starndard/advance level
                new_level_group, # new level group
                row[5].title(), # student first Name
                row[6].title(), # student middle Name
                row[7].title(), # student last Name
                row[8].title(), # student date of birh
                row[9],  # student gender
                row[10].title(), # student regular school
                row[11].title(), # school year
                row[14].title(), # Father's first name
                row[15].title(), # Father's last name
                row[16], # Father's mobile
                row[17], # Father's email address
                row[18].title(), # Mother's first name
                row[19].title(), # Mother's last name
                row[20], # Mother's mobile
                row[21], # Mother's email address
                row[22].title(), # address: street
                row[24].title(), # address: suburb
                row[25].upper(), # address: state
                row[26], # address: post code
                row[28], # any allergies
                row[29], # medicare card
                row[30].title(), # doctor/clinic
                row[31], # doctor/clinic phone
                row[32].title(), # emergency contact 1 (first name)
                row[33].title(), # emergency contact 1 (last name)
                row[34], # emergency contact 1 (phone)
                row[35].title(), # emergency contact 1 (relation)
                row[36].title(), # emergency contact 2 (first name)
                row[37].title(), # emergency contact 2 (last name)
                row[38], # emergency contact 2 (phone)
                row[39].title(), # emergency contact 2 (relation)
                row[17] + ';' + row[21], # recipient (combined email address parents)
                '', # email sent flag
                GNPS_ID  + GNPS_DOMAIN if row[3]=='No' else row[4] + GNPS_DOMAIN,
                pwd_str if row[3]!='Yes' else 'Use Existing Password', # random password
                GNPS_ORG_PREFIX + str(row[2]).upper() + '/' + str(level_group.rsplit("_")[0] + '-' + level_group.rsplit("_")[1]).upper() if row[3]=='No' else GNPS_ORG_PREFIX + str(row[2]).upper() + '/' + str(group_info.rsplit("@")[0].rsplit("_")[0]).upper() + '-' + str(group_info.rsplit("@")[0].rsplit("_")[1]).upper() if group_info !='' else '', # Org Path
                ])
        else:
            # if res != None :
            #     print(str(res['primaryEmail']))
            f2f_records.append([
                calculate_age(row[8]),
                row[3], # new or existing student
                row[2], # class format online/face-to-face
                row[48] if row[48] !='' else 'n/a', # reason for selecting online
                GNPS_ID if row[3]=='No' else row[4], # if existing student, list student ID
                row[47], # Invoice ID
                'Session-1' if row[3]=='No' else 'Session-2', # session
                F2F_CLASS_TIMINGS_SESSION1 if row[3]=='No' else F2F_CLASS_TIMINGS_SESSION2, #class time
                '', # room
                '', # class code
                '', # class id
                '', # class name
                '', # Teachers
                'LEVEL-BASIC' if row[3]=='No' else str(res['orgUnitPath']).rsplit("/")[len(str(res['orgUnitPath']).rsplit("/"))-1] if res != None else '', # check and insert level from the system
                level_group + GNPS_DOMAIN if row[3]=='No' else group_info,
                new_level, # new level
                level_type, # starndard/advance level
                new_level_group, # new level group
                row[5].title(), # student first Name
                row[6].title(), # student middle Name
                row[7].title(), # student last Name
                row[8].title(), # student date of birh
                row[9],  # student gender
                row[10].title(), # student regular school
                row[11].title(), # school year
                row[14].title(), # Father's first name
                row[15].title(), # Father's last name
                row[16], # Father's mobile
                row[17], # Father's email address
                row[18].title(), # Mother's first name
                row[19].title(), # Mother's last name
                row[20], # Mother's mobile
                row[21], # Mother's email address
                row[22].title(), # address: street
                row[24].title(), # address: suburb
                row[25].upper(), # address: state
                row[26], # address: post code
                row[28], # any allergies
                row[29], # medicare card
                row[30].title(), # doctor/clinic
                row[31], # doctor/clinic phone
                row[32].title(), # emergency contact 1 (first name)
                row[33].title(), # emergency contact 1 (last name)
                row[34], # emergency contact 1 (phone)
                row[35].title(), # emergency contact 1 (relation)
                row[36].title(), # emergency contact 2 (first name)
                row[37].title(), # emergency contact 2 (last name)
                row[38], # emergency contact 2 (phone)
                row[39].title(), # emergency contact 2 (relation)
                row[17] + ';' + row[21], # recipient (combined email address parents)
                '', # email sent flag
                GNPS_ID  + GNPS_DOMAIN if row[3]=='No' else row[4] + GNPS_DOMAIN,
                pwd_str if row[3]!='Yes' else 'Use Existing Password', # random password
                GNPS_ORG_PREFIX + str(row[2]).upper() + '/' + str(level_group.rsplit("_")[0] + '-' + level_group.rsplit("_")[1]).upper() if row[3]=='No' else GNPS_ORG_PREFIX + str(row[2]).upper() + '/' + str(group_info.rsplit("@")[0].rsplit("_")[0]).upper() + '-' + str(group_info.rsplit("@")[0].rsplit("_")[1]).upper() if group_info !='' else '', # Org Path

                ])
#
online_resource = {
    "majorDimension": "ROWS",
    "values": online_records
}
#
f2f_resource = {
    "majorDimension": "ROWS",
    "values": f2f_records
}
#
sheet = spreadsheet_service.spreadsheets()
# process online records
sheet.values().append(
    spreadsheetId=PROCESSED_ENROLMENTS_SPREADSHEET_ID,
    range=PROCESSED_ONLINE_ENROLMENTS_RANGE_NAME,
    valueInputOption='USER_ENTERED',
    insertDataOption = 'OVERWRITE',
    body=online_resource
).execute()

# process face-to-face records
sheet.values().append(
    spreadsheetId=PROCESSED_ENROLMENTS_SPREADSHEET_ID,
    range=PROCESSED_F2F_ENROLMENTS_RANGE_NAME,
    valueInputOption='USER_ENTERED',
    insertDataOption = 'OVERWRITE',
    body=f2f_resource
).execute()