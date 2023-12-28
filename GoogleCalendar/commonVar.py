class commonVar:
    def __init__(self, fullSche = None, lecNameList = None, lecList = None, firstMon = None):
        self.fullSche = fullSche
        self.lecNameList = lecNameList
        self.lecList = lecList
        self.firstMon = firstMon
        
    def Set_fullSche(self, fullSche):
        self.fullSche = fullSche
    
    def Set_lecNameList(self, lecNameList):
        self.lecNameList = lecNameList
        
    def Set_firstMon(self, firstMon):
        self.firstMon = firstMon        
        
    def Set_lecList(self, lecList):
        self.lecList = lecList
    
    def get_fullSche(self):
        return self.fullSche
    
    def get_lecNameList(self):
        return self.lecNameList
    
    def get_lecList(self):
        return self.lecList

    def get_firstMon(self):
        return self.firstMon    
     
common = commonVar(None, None, None)