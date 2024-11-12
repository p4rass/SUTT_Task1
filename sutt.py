import pandas as pd
import json
import logging

logging.basicConfig(level=logging.DEBUG)

def handle_error(error_type, message, context=""):
    logging.error(f"{context} - {error_type}: {message}")
    return 'Not Found'

path = r'C:\Users\paras\OneDrive\Desktop\SUTT\timetable.xlsx'

def load(name, df, line):
    try:
        if name not in df.columns:
            raise KeyError(f"Missing expected column: {name}")
        
        if len(df) <= line:
            raise IndexError(f"Row index {line} is out of bounds.")
        return df[name].iloc[line]
    except (KeyError, IndexError, ValueError) as e:
        return handle_error(type(e), str(e), context=f"Loading data for column '{name}' at row {line}")
    except Exception as e:
        return handle_error(type(e), f"Unexpected error occurred: {e}", context=f"Loading data for column '{name}' at row {line}")

for i in range(6):
    print(" -------------------------------------------------------------------------------------------------------------------------------------------------------------------- ")
    print(" -------------------------------------------------------------------------------------------------------------------------------------------------------------------- ")

    try:
        df1 = pd.read_excel(path, sheet_name=i, header=1)
        df2 = pd.read_excel(path, sheet_name=i, header=2)
    except FileNotFoundError:
        logging.error(f"File not found: {path}", context=f"Reading sheet {i}")
        continue
    except PermissionError:
        logging.error(f"Permission denied when trying to open: {path}", context=f"Reading sheet {i}")
        continue
    except Exception as e:
        logging.exception(f"Unexpected error while reading Excel file: {e}", context=f"Reading sheet {i}")
        continue

    course_code = load('COM COD', df1, 1) if isinstance(df1['COM COD'].iloc[1], (float, int)) else 'Not Found'
    course_title = load('COURSE TITLE', df1, 1)
    course_number = load('COURSE NO.', df1, 1)
    credit_str = {
        'Lecture': float(load('L', df2, 0)) if isinstance(load('L', df2, 0), (float, int)) else 'Not Found',
        'Practical': float(load('P', df2, 0)) if isinstance(load('P', df2, 0), (float, int)) else 'Not Found',
        'Units': float(load('U', df2, 0)) if isinstance(load('U', df2, 0), (float, int)) else 'Not Found'
    }

    data = {
        "Course Code : ": course_code,
        "Course Title : ": course_title,
        "Course Number : ": course_number,
        "Credits : ": credit_str
    }

    class Section:
        def __init__(self, sec_type, sec_num, instructor, room_no, time_slot):
            self.sec_type = sec_type
            self.sec_num = sec_num
            self.instructor = instructor
            self.room_no = room_no
            self.time_slot = time_slot
            self.Instructors = [instructor]
            self.timings(time_slot)

        def timings(self, time_slot):
            self.day = []
            self.day1 = []
            self.time = []
            self.time1 = []
            if time_slot:
                split_time = time_slot.split()
            else:
                logging.error(f"Missing or malformed time slot: {time_slot}", context=f"Processing section {self.sec_num}")
                return

            d = 0
            e = 0
            f = 0
            g = 0
            for i in range(len(split_time)):
                if i<= (len(split_time)-2):
                    if split_time[i].isdigit() and split_time[i+1].isdigit():
                        if int(split_time[i])+1 == int(split_time[i+1]):
                            g = 1
                            if int(split_time[i]) <= 2:
                                split_time[i] = str(int(split_time[i])+7) + "AM  to " + str(int(split_time[i])+9) +"AM"
                            elif int(split_time[i]) == 3:
                                split_time[i] = "10 AM to 12 PM"
                            elif int(split_time[i]) == 4:
                                split_time[i] = "11 AM to 1 PM"
                            elif int(split_time[i]) >= 5:
                                split_time[i] = str(int(split_time[i])-5) + "PM  to " + str(int(split_time[i])-3) +"PM"
                            self.time.append(split_time[i])
                            d = 1
                        else:
                            pass
                    else:
                        pass

                if g==0:
                    if d == 0 and not split_time[i].isdigit():
                        if split_time[i] == "M":
                            self.day.append("Monday")
                        if split_time[i] == "T":
                            self.day.append("Tuesday")
                        if split_time[i] == "W":
                            self.day.append("Wednesday")
                        if split_time[i] == "Th":
                            self.day.append("Thursday")
                        if split_time[i] == "F":
                            self.day.append("Friday")
                        if split_time[i] == "S":
                            self.day.append("Saturday")
                    elif split_time[i].isdigit() and e == 0 and f == 0:
                        if int(split_time[i]) <= 3:
                            split_time[i] = str(int(split_time[i])+7) + "AM  to " + str(int(split_time[i])+8) +"AM"
                        elif int(split_time[i]) == 4:
                            split_time[i] = str(int(split_time[i])+7) + "AM  to " + str(int(split_time[i])+8) +"PM"
                        elif int(split_time[i]) == 5:
                            split_time[i] = str(int(split_time[i])+7) + "PM  to " + str(int(split_time[i])-4) +"PM"
                        elif int(split_time[i]) >= 6:
                            split_time[i] = str(int(split_time[i])-5) + "PM  to " + str(int(split_time[i])-4) +"PM"
                        self.time.append(split_time[i])
                        d = 1
                        f = 1

                    elif d == 1 and not split_time[i].isdigit():
                        e = 1
                        if split_time[i] == "M":
                            self.day1.append("Monday")
                        if split_time[i] == "T":
                            self.day1.append("Tuesday")
                        if split_time[i] == "W":
                            self.day1.append("Wednesday")
                        if split_time[i] == "Th":
                            self.day1.append("Thursday")
                        if split_time[i] == "F":
                            self.day1.append("Friday")
                        if split_time[i] == "S":
                            self.day1.append("Saturday")
                    elif split_time[i].isdigit() and e == 1:
                        if int(split_time[i]) <= 3:
                            split_time[i] = str(int(split_time[i])+7) + "AM  to " + str(int(split_time[i])+8) +"AM"
                        elif int(split_time[i]) == 4:
                            split_time[i] = str(int(split_time[i])+7) + "AM  to " + str(int(split_time[i])+8) +"PM"
                        elif int(split_time[i]) == 5:
                            split_time[i] = str(int(split_time[i])+7) + "PM  to " + str(int(split_time[i])-4) +"PM"
                        elif int(split_time[i]) >= 6:
                            split_time[i] = str(int(split_time[i])-5) + "PM  to " + str(int(split_time[i])-4) +"PM"
                        self.time1.append(split_time[i])

                time_str = ', '.join(self.time)
                time1_str = ', '.join(self.time1)
                self.slot = {**{item1: time_str for item1 in self.day}, **{item2: time1_str for item2 in self.day1}}

        def dict(self):
            return {
                "Section Type : ": self.sec_type,
                "Section Number :": self.sec_num,
                "Room No : ": self.room_no,
                "Instructors : ": self.Instructors,
                "Timing : ": self.slot
            }

        def add_inst(self, instructor):
            if instructor not in self.Instructors:
                self.Instructors.append(instructor)

    Sections = []
    last_section = None
    lecturz = "lecture"
    for i in range(df1.shape[0]):
        sec_num = load('SEC', df1, i)
        instructor = load('INSTRUCTOR-IN-CHARGE / Instructor', df1, i)
        room_no = load('ROOM', df1, i)
        time_slot = load('DAYS & HOURS', df1, i)
        lecture = load('COURSE TITLE', df1, i)

        if lecture == 'Tutorial':
            lecturz = "Tutorial"
        elif lecture == 'Practical':
            lecturz = "Practical"

        if pd.notna(sec_num):
            section = Section(
                lecturz,
                sec_num,
                instructor,
                room_no,
                time_slot
            )
            Sections.append(section.dict())
            last_section = section
        elif last_section:
            last_section.add_inst(instructor)

    merged_data = {**data, "Sections : ": Sections}
    try:
        jz = json.dumps(merged_data, default=str, indent=4)
    except TypeError:
        logging.error("Failed to serialize data to JSON. Ensure all data is serializable.", context="Merging data for JSON output")
    print(jz)
