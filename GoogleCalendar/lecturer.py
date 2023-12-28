class lecturer:
    def __init__(self, lec_name):
        self.Name = lec_name
        self.Schedule = []
        
    def scheAdd(lecList, fullSche):
        for lec in lecList:
            for element in fullSche:
                if lec.Name == element._lec:
                    lec.Schedule.append(element)
    
    def getLecNameList(fullsche):
        seen = set()
        lecNameList = []
        for e in fullsche:
            if e._lec not in seen:
                seen.add(e._lec)
                lecNameList.append(e._lec)
        return lecNameList

    def lecCre(lecNameList):      
        LecList = []
        for name in lecNameList:
            Lecturer = lecturer(name)
            LecList.append(Lecturer)
        return LecList
    
    def getEventData(i, j, lecList):        
        eventTitle = lecList[i].Schedule[j]._course_name + ' - ' +lecList[i].Schedule[j]._class_name
        day = lecList[i].Schedule[j]._day
        
        if lecList[i].Schedule[j]._room != None:
            room = lecList[i].Schedule[j]._room
        else:
            room = ''

        week = lecList[i].Schedule[j]._week
        lecName = lecList[i].Schedule[j]._lec
        sessionDict = {'1-3': ('T06:45:00.000+07:00', 'T09:10:00.000+07:00'),
            '4-6': ('T09:25:00.000+07:00', 'T11:50:00.000+07:00'),
            '7-9':('T12:15:00.000+07:00', 'T14:40:00.000+07:00'),
            '10-12':('T14:55:00.000+07:00', 'T17:20:00.000+07:00')
            }
        session = sessionDict[lecList[i].Schedule[j]._session]
        
        return eventTitle, room, day, week, session, '', '', lecName