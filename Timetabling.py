import os
import pandas as pd
from math import ceil

SKIP = ['PE001', 'CS5030', 'CS5000']

# Create class Room
class Room:
    def __init__(self, roomName, roomCapacity, roomType):
        self.roomName = roomName
        self.roomCapcity = roomCapacity
        self.roomType = roomType
        self.slot = {day:[0]*10 for day in ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday']}

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
    sortedData = data.sort_values(by=['Type', 'Capacity'], ignore_index=True)
    roomList = []

    for i in sortedData.index:
        room = Room(sortedData['RoomName'][i], sortedData['Capacity'][i], sortedData['Type'][i])
        roomList.append(room)
    
    return roomList

# Function: load class list from excel to python object list
def loadClassList(classListFile):
    data = pd.read_excel(classListFile, sheet_name=0)
    data["classPerSubject"] = ""
    
    for i in data.index:
        data['classPerSubject'][i] = len(data[data.subjectID == data['subjectID'][i]])
    sortedData = data.sort_values(by=['studentYear','classPerSubject', 'classCredits', 'classID'], ascending=[True, False, False, True], ignore_index=True)
    print(sortedData)
    classList = []

    for i in sortedData.index:
        inputClass = Class(
        sortedData['subjectID'][i],
        sortedData['classID'][i],
        sortedData['className'][i],
        sortedData['lecturerID'][i],
        sortedData['lecturerName'][i],
        sortedData['classCapacity'][i],
        sortedData['classCredits'][i],
        sortedData['practice'][i],
        sortedData['classType'][i],
        sortedData['skipWeek'][i],
        sortedData['studentYear'][i],
        sortedData['term'][i],
        sortedData['schoolYear'][i],
        sortedData['program'][i],
        sortedData['faculty'][i],
        sortedData['startDate'][i],
        sortedData['endDate'][i],
        sortedData['language'][i]
        )
        classList.append(inputClass)
    
    return classList

def exportClassList(classList, classListFile):
    data = pd.read_excel(classListFile, sheet_name=0)
    data["classPerSubject"] = ""
    
    for i in data.index:
        data['classPerSubject'][i] = len(data[data.subjectID == data['subjectID'][i]])
    sortedData = data.sort_values(by=['studentYear','classPerSubject', 'classCredits', 'classID'], ascending=['true', 'false', 'false', 'true'], ignore_index=True)
    
    for i in sortedData.index:
        sortedData['day'][i] = classList[i].day
        sortedData['lesson'][i] = classList[i].lesson
        sortedData['classRoom'][i] = classList[i].classRoom

    sortedData.to_excel("Output.xlsx")

def putClassToTimetable(inputClass, timetable, roomList, classList):
    index = roomList.index(next(x for x in roomList if x.roomName == inputClass.classRoom))
    beginLesson = index * 10 + int(inputClass.lesson[0])
    endLesson = beginLesson + len(inputClass.lesson) - 1
    for i in range(beginLesson, endLesson+1):
        timetable[inputClass.day][i] = inputClass.classID

def findRoom(roomList, timetable, classID, roomType, classCapacity, classCredits, previousDay):
    # Check each room in room list
    for room in roomList:
        # Check room based on class condition
        if room.roomType == roomType and room.roomCapcity >= classCapacity:
            # Check each day in a week
            for day in ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday']:
                """ CONDITION: classes of same student year are not in the same day """
                if day == previousDay:
                    continue
                # Find free slot
                freeSlot = [i for i, x in enumerate(room.slot[day]) if x == 0]
                # For each free slot, check availabity of free period
                for slot in freeSlot:
                    # The free period must between lesson 1 -> 5 and lesson 6 -> 10
                    if (slot <= 4 - classCredits) or (4 < slot and slot <= 9 - classCredits):
                        # If begining slot and endding slot are free, ALL slots between them are free
                        if room.slot[day][slot + classCredits] == 0:
                            lesson = ""
                            # Generate available lesson from free slot
                            for i in range(slot + 1, slot + classCredits + 1):
                                room.slot[day][i - 1] = classID
                                lesson += str(i)
                            return room.roomName, day, lesson
    # If no available free slot found, return ERROR
    return ValueError

def init(roomList, classList):
    # Import room list
    roomList = loadRoomList(roomList)
    # Import class list
    classList = loadClassList(classList)
    # Init empty timetable
    d = {day:[0]*len(roomList)*10 for day in ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday']}
    timetable = pd.DataFrame(data=d, index=list(range(1,len(roomList)*10+1)))
    timetable['roomName'] = ""
    for i in range(1,len(roomList)*10):
        timetable['roomName'][i] = roomList[ceil(i/10) - 1].roomName
    timetable['roomName'][1160] = roomList[-1].roomName
    
    return roomList, classList, timetable

def main():
    roomList, classList, timetable = init("RoomList.xlsx", "Input.xlsx")
    previousDay = "Sunday"

    for eachClass in classList:
        if eachClass.subjectID in SKIP:
            continue

        classID = eachClass.classID
        classCapacity = eachClass.classCapacity
        classCredits = eachClass.classCredits
        if eachClass.program != "CQUI":
            roomType = "CLC"
        else:
            roomType = "CQUI"

        roomName, day, lesson = findRoom(roomList, timetable, classID, roomType, classCapacity, classCredits, previousDay)
        eachClass.classRoom = roomName
        eachClass.day = day
        eachClass.lesson = lesson
        previousDay = day
        
        putClassToTimetable(eachClass, timetable, roomList, classList)

    timetable = timetable.set_index("roomName", append=True).swaplevel(0, 1)

    timetable.to_excel("Output_Timetable.xlsx")

    exportClassList(classList, "Input.xlsx")

if __name__ == "__main__":
    main()