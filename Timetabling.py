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
    roomList = []

    for i in data.index:
        room = Room(data['RoomName'][i], data['Capacity'][i], data['Type'][i])
        roomList.append(room)
    
    return roomList

# Function: load class list from excel to python object list
def loadClassList(classListFile):
    data = pd.read_excel(classListFile, sheet_name=0)
    classList = []

    for i in data.index:
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

def putClassToTimetable(inputClass, timetable, day, lesson, room, roomList, classList):
    index = roomList.index(next(x for x in roomList if x.roomName == room))
    beginLesson = index * 10 + int(lesson[0])
    endLesson = beginLesson + len(lesson)
    for i in range(beginLesson, endLesson+1):
        timetable[day][i] = inputClass

    print("Class: " + inputClass)
    print("Room: " + room)
    print("Index in list: " + str(index))
    print("Begin lesson: " + str(beginLesson))
    print("End lesson: " + str(endLesson))
    print(timetable)

roomList = loadRoomList("RoomList.xlsx")
classList = loadClassList("Input.xlsx")

d = {day:[0]*len(roomList)*10 for day in list(range(2,8,1))}
timetable = pd.DataFrame(data=d, index=list(range(1,len(roomList)*10+1)))

putClassToTimetable("IT007.L11", timetable, 2, "12345", "A205", roomList, classList)