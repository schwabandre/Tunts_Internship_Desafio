from __future__ import print_function
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from math import ceil

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# The ID and range of a sample spreadsheet.
SPREADSHEET_ID = '1i5KHkTwDMYouxyhgQ-JetMT8LGU_IdqZ6u3iLYLNK9o'
SPREADSHEET_NAME = 'engenharia_de_software'
TOTAL_CLASSES_NUMBER_RANGE = 'A2'
STUDENTS_DATA_RANGE = 'A4:F'

# Student Situation
FAILED_BY_GRADE = "Reprovado por Nota"
FAILED_BY_UNPRESENCE = "Reprovado por Falta"
APPROVED = "Aprovado"
FINAL_EXAM = "Exame Final"

def main():
    print('***************************************')
    print('****Welcome to Grade-Calculator****')
    print('***************************************')
    print('Get Credentials')
    creds = None
    # The file token.pickle stores the authentication token to access the spreadsheet.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('sheets', 'v4', credentials=creds)

    # Requesting Google API for total Classes Number
    print('Requesting for total Classes Number')
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=SPREADSHEET_ID,
                                range=TOTAL_CLASSES_NUMBER_RANGE).execute()
    classes = result.get('values', [])

    if not classes:
        print('Could not fetch spreadsheet data')
        return
        
    total_classes = int(classes[0][0].split(':')[1])
    maximum_missed_classes = total_classes * 0.25

    print('Requesting Students Data')
    data_result = sheet.values().get(spreadsheetId=SPREADSHEET_ID,
                                range=STUDENTS_DATA_RANGE).execute()
    students_data = data_result.get('values', [])
    
    if not students_data:
        print('Could not fetch spreadsheet data')
        return

    print('Data fetched Successfully')

#Students grades calculation
    print('Calculating Students Situation')
    students_conditions = []
    for student in students_data:
        [id, name, missed_classes, p1, p2, p3] = student
        student_situation = ''
        final_grade = 0
        grade = float(int(p1) + int(p2) + int(p3))/3 
        
        #Determination for Pass or Fail Students
        if  int(missed_classes) > maximum_missed_classes:
            student_situation = FAILED_BY_UNPRESENCE
            print('{} Failed due to {} absent classes.'.format(name, missed_classes))    
        
        elif grade < 50:
            student_situation = FAILED_BY_GRADE
            print('{} Fail by Grade, with {:.2f} score'.format(name, grade))
        
        elif 50 <= grade < 70:
            student_situation = FINAL_EXAM
            final_grade = ceil(100 - grade)
            print('{} got a score of {:.2f}, and is requested to do Final Exam, with a minimum score of {} to be approved'.format(name, grade, final_grade))
            
        else:
            student_situation = APPROVED
            print('{} is approved with a {:.2f} score'.format(name, grade))

        students_conditions.append( [student_situation, final_grade] )

#Definition of what will be shown in the spreadsheet
    students_conditions_content = {'values': students_conditions}
    last_student_row = 3+ len(students_data)
    students_conditions_range = 'G4:H{}'.format(last_student_row)

    print('Writing results in the spreadsheet')
    result = sheet.values().update(spreadsheetId=SPREADSHEET_ID, range=students_conditions_range, valueInputOption = 'RAW',
       body=students_conditions_content).execute()    
            
#End of the program Grade-Calculator
    print('Writing was succesfully executed')
if __name__ == '__main__':
    main()