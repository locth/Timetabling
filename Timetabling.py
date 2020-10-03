import os
import pandas as pd

# Create class Room
class Room:
    def __init__(self, roomName, roomCapacity, roomType):
        self.roomName = roomName
        self.roomCapcity = roomCapacity
        self.roomType = roomType

# Create class Class
class Class:
    def __init__(self, subjectID, classID, className, lecturerID, lecturerName, classCapacity, classCredits, practice, classType, skipWeek, studentYear, term, schoolYear, program, faculty, startDate, endDate, language):
        self.subjectID = subjectID
        self.classID = classID
        self.className = className
        self.lecturerID = lecturerID
        self.lectuterName = lecturerName
        self.classCapacity = classCapacity
        self.classCredits = classCredits
        self.practice = practice
        self.classType = classType
        self.day = ""
        self.lesson = ""
        self.skipWeek = skipWeek
        self.classRoom = ""
        self.studentYear = studentYear
        self.term = term
        self.schoolYear = schoolYear
        self.program = program
        self.faculty = faculty
        self.startDate = startDate
        self.endDate = endDate
        self.language = language

# Function: load room list from excel to python object list
def loadRoomList(roomListFile):
    data = pd.read_excel(roomListFile, sheet_name=0)
    # SORT data by room's type and room's capacity
    sortedData = data.sort_values(by=['Type', 'Capacity'])
    roomList = []

    for i in data.index:
        room = Room(data['RoomName'][i], data['Capacity'][i], data['Type'][i])
        roomList.append(room)
    
    return roomList

# Function: load class list from excel to python object list
def loadClassList(classListFile):
    data = pd.read_excel(classListFile, sheet_name=0)
    data["classPerSubject"] = 0
    
    for i in data.index:
        data['classPerSubject'][i] = len(data[data.subjectID == data['subjectID'][i]])
    sortedData = data.sort_values(by=['studentYear','classPerSubject', 'classCredits', 'classID'])
    classList = []

    for i in sortedData.index:
        inputClass = Class(
        data['subjectID'][i],
        data['classID'][i],
        data['className'][i],
        data['lecturerID'][i],
        data['lecturerName'][i],
        data['classCapacity'][i],
        data['classCredits'][i],
        data['practice'][i],
        data['classType'][i],
        data['skipWeek'][i],
        data['studentYear'][i],
        data['term'][i],
        data['schoolYear'][i],
        data['program'][i],
        data['faculty'][i],
        data['startDate'][i],
        data['endDate'][i],
        data['language'][i]
        )
        classList.append(inputClass)
    
    return classList

def putClassToTimetable(inputClass, timetable, roomList, classList):
    index = roomList.index(next(x for x in roomList if x.roomName == inputClass.classRoom))
    beginLesson = index * 10 + int(inputClass.lesson[0])
    endLesson = beginLesson + len(inputClass.lesson) - 1
    for i in range(beginLesson, endLesson+1):
        timetable[inputClass.day][i] = inputClass.classID

def init(roomList, classList):
    # Import room list
    roomList = loadRoomList(roomList)
    # Import class list
    classList = loadClassList(classList)
    # Init empty timetable
    d = {day:[0]*len(roomList)*10 for day in ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday']}
    timetable = pd.DataFrame(data=d, index=list(range(1,len(roomList)*10+1)))
    return roomList, classList, timetable

roomList, classList, timetable = init("RoomList.xlsx", "Input.xlsx")

for eachClass in classList:
    print(eachClass.classID)
 
print(timetable)