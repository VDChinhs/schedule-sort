from __future__ import print_function
import datetime
import os.path
import json

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

class CalendarFunc:
    SCOPES = ['https://www.googleapis.com/auth/calendar']
    
    def credsCheck():
        creds = None
        if os.path.exists('Cred/token.json'):
            creds = Credentials.from_authorized_user_file('Cred/token.json', CalendarFunc.SCOPES)
        # If there are no (valid) credentials available, let the user log in.
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    'Cred/credentials.json', CalendarFunc.SCOPES)
                creds = flow.run_local_server(port=0)
            # Save the credentials for the next run
            with open('Cred/token.json', 'w') as token:
                token.write(creds.to_json())
        return creds
    
    def buildService():
        creds = CalendarFunc.credsCheck()
        service = build('calendar', 'v3', credentials=creds)
        return service
    
    def credRemove():
        if os.path.exists('Cred/token.json'):
            os.remove('Cred/token.json')
            return 0
        else:
            return 1
    
    def IDFileRemove():
        if os.path.exists('data/CalendarID.json'):
            os.remove('data/CalendarID.json')
            return 0
        else:
            return 1
        
    def CreateIDFile():
        filepath = 'data/CalendarID.json'
        ID = {}
        with open(filepath, "w") as json_file:
            json.dump(ID, json_file)
        return
    
    def writeCaleID(dict):
        filepath = 'data/CalendarID.json'
        with open (filepath, 'w') as file:
                json.dump(dict, file)        
        return 0
                    
    def readCaleID():
        if os.path.exists('data/CalendarID.json'):
            filePath = 'data/CalendarID.json'
            with open(filePath, 'r') as jsonFile:
                IDDict = json.load(jsonFile)
            return IDDict
        else:
            return 1
    
    def newCalendar(lec):
        try:
            service = CalendarFunc.buildService()
            calendar = {
                'summary': '',
                'timeZone': 'Asia/Ho_Chi_Minh'
            }
            calendar['summary'] = lec
        
            created_calendar = service.calendars().insert(body=calendar).execute() 
            return created_calendar['id']
        
        except HttpError as error:
            return error
        
    def eventCreate(summary, room, day, week, sess, sDate, eDate, lecName):
        service = service = CalendarFunc.buildService()
        event = {
        'summary': 'CourseTitle - ClassName',
        'location': 'Room',
        'start': {
            'dateTime': 'yyyy-mm-ddThh:mm:ss.000+07:00',
            'timeZone': 'Asia/Ho_Chi_Minh'
        },
        'end': {
            'dateTime': 'yyyy-mm-ddThh:mm:ss.000+07:00',
            'timeZone': 'Asia/Ho_Chi_Minh'
        },
        'recurrence': [
            'RRULE:FREQ=WEEKLY;UNTIL=yyyymmddThhmmssZ',
        ],
        }
        
        event['summary'] = summary
        event['location'] = room
        sDateTime =  sDate + sess[0]
        event['start']['dateTime'] = sDateTime
        eDateTime = sDate + sess[1]
        event['end']['dateTime'] = eDateTime
        rRuleStr = 'RRULE:FREQ=WEEKLY;UNTIL=' + eDate + 'T170000Z'
        event['recurrence'][0] = rRuleStr
        
        IDDict = CalendarFunc.get_calendar_id()
        ID = IDDict[lecName]
        
        try:
            recurring_event = service.events().insert(calendarId=ID, body=event).execute()
            return recurring_event['id']
        
        except HttpError as error:
            return 'An error occurred: %s' % error
    
    def calendarRemove(caleID):
        service = CalendarFunc.buildService()
        try:    
            service.calendars().delete(calendarId=caleID).execute()
            return caleID
        except HttpError as error:
            return error
        
    # def getFirstMonday():
    #     # Check if saved date exists
    #     if os.path.exists('data/firstMon'):
    #         with open ('data/firstMon', 'r') as file:
    #             dateread = file.readline()
    #             if dateread:
    #                 firstMonDate = datetime.datetime.strptime(dateread, "%Y-%m-%d").date()
    #                 return firstMonDate
    #             # If exists file is empty, delete file
    #             else:
    #                 file.close()
    #                 os.remove('data/firstMon')
    #                 CalendarFunc.getFirstMonday()
    #     # In case date file not exist
    #     else:
    #         date_input = input("Enter a date (DD-MM-YYYY): ")
    #         try:
    #             firstMonDate = datetime.datetime.strptime(date_input, "%d-%m-%Y").date()
    #             with open ('data/firstMon', 'w') as file:
    #                 file.write(str(firstMonDate))
    #                 return firstMonDate
    #         except ValueError:
    #             print("Invalid datetime format. Please use the format DD-MM-YYYY")
    #             CalendarFunc.getFirstMonday()
        
    def getStages(day, weekStr, firstMon):
        semesterSD = firstMon
        # Standardization of input string
        weeks = [' ']*22
        list1 = [int for int in weekStr]
        l = len(list1)
        for i in range(l):
            if weeks[i] != list1[i]:
                weeks[i] = list1[i]
        
        # Create variables to keep track of date      
        stage = []
        work_stage = 0
        startDate = semesterSD + datetime.timedelta(days=(day-2))
        current_date = startDate
        
        for index, week in enumerate(weeks):
            if week != ' ':
                work_stage += 1
            else:
                if work_stage !=0:
                    stageED = current_date + datetime.timedelta(weeks=work_stage-1)
                    stage.append((current_date, stageED))
                    work_stage = 0
                # update current_date
                
                current_date = startDate + datetime.timedelta(weeks=index+1)
                
        #handle the last stage end with working week
        if work_stage != 0:
            endDate = current_date + datetime.timedelta(weeks=work_stage-1)
            stage.append((current_date, endDate))
        return stage
    
    def invite(calendarID, email):
        service = CalendarFunc.buildService()
        access_rule = {
        'scope': {
            'type': 'group',
            'value': 'scopeEmail',
            },
        'role': 'writer'
        }
        access_rule['scope']['value'] = email
        
        try:
            service.acl().insert(calendarId=calendarID, body=access_rule).execute()
            return 'successful'
        
        except HttpError as error:
            return error
    
    def get_calendar_list():
        service = CalendarFunc.buildService()
        calendar_list = service.calendarList().list().execute()
        return calendar_list
    
    def get_calendar_id():
        calendar_list = CalendarFunc.get_calendar_list()
        id_dict = {}
        for calendar_list_entry in calendar_list['items']:
            id_dict.update({calendar_list_entry['summary']:calendar_list_entry['id']})
        return id_dict       