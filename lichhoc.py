class monhoc:
    def __init__(self,STT,COURSE_CODE,COURSE_NAME,CLASS_NAME,CLASS_COM,DAY,SESSION,ROOM,CREDIT,START,END,WEEK,LEC=None):
        self._stt = STT
        self._course_code = COURSE_CODE
        self._course_name = COURSE_NAME
        self._class_name = CLASS_NAME
        self._class_com = CLASS_COM
        self._day = DAY
        self._session = SESSION
        self._room = ROOM
        self._credit = CREDIT
        self._start = START
        self._end = END
        self._week = WEEK
        self._lec = LEC
        self._lopghep = None
        self._trung = None
        self._tuantrung = []