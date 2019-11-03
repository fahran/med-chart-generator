from datetime import datetime
from datetime import date
from datetime import timedelta
from sys import stdin
import xlsxwriter

class Medication:
    days = 7
    dose = None

    def __init__(self, name: str, interval: timedelta, start_time: datetime, dose=None):
        self.name = name
        self.dose = dose
        self.interval = interval
        self.start_time = start_time
        self.timetable = self.create_timetable()

    def create_timetable(self):
        timestamp = self.start_time
        timestamps = []
        while timestamp > datetime.today().replace(hour=0, minute=0, second=0, microsecond=0):
            timestamps.append(timestamp)
            timestamp = timestamp - self.interval

        timestamps.reverse()
        timestamp = self.start_time

        while timestamp < self.start_time + timedelta(days=7):
            timestamp = timestamp + self.interval
            timestamps.append(timestamp)
        return timestamps

    def __str__(self):
        return "Name: %s, Interval: %s, Start Time: %s, Dose: %s" % (self.name, self.interval, self.start_time, self.dose)

def main():
    # medications = __static_data()
    medications = __ask_questions()
    for medication in medications:
        print(medication)
    __produce_spreadsheet(medications)

def __static_data():
    one = Medication("Malteasers", timedelta(hours=3, minutes=45), datetime(2019, 11, 3, 14, 15), 1.5)
    two = Medication("Candy Canes", interval=timedelta(hours=6), start_time=datetime(2019, 11, 3, 7))
    three = Medication("Plums", interval=timedelta(hours=6), start_time=datetime(2019, 11, 3, 9))
    four = Medication("Fudge", interval=timedelta(hours=24), start_time=datetime(2019, 11, 3, 8), dose=1)
    five = Medication("Revels", interval=timedelta(hours=24), start_time=datetime(2019, 11, 3, 18), dose=1)
    return [one, two, three, four, five]

def __ask_questions():
    print("Hello! How many types of tablet do you want to track?")
    medicationCount = int(stdin.readline())
    medications = []
    for i in range(0, medicationCount):
        print("What's the name of the next medication?")
        name = stdin.readline().rstrip()
        print("How often do you take %s? (eg. 3h45" % name)
        intervalSegs = stdin.readline().rstrip().split("h")
        minutes = int(intervalSegs[1]) if intervalSegs[1] != "" else 0
        interval = timedelta(hours=int(intervalSegs[0]), minutes=minutes)
        print("When was the last time you took %s? eg. 16:00" % name)
        timeSegs = stdin.readline().rstrip().split(":")
        start_time = datetime.now().replace(hour=int(timeSegs[0]), minute=int(timeSegs[1]))
        print("How much do you take? Just press 'Enter' if you don't want to track it.")
        dose = None
        inputDose = stdin.readline().rstrip()
        if (inputDose != ""):
            dose = inputDose

        medications.append(Medication(name, interval, start_time, dose))

    for medication in medications:
        print(medication)

    return medications


def __produce_spreadsheet(medications):
    workbook = xlsxwriter.Workbook('medications.xlsx')
    worksheet = workbook.add_worksheet()
    normal_cell = workbook.add_format({'align': 'center', 'valign': 'center', 'border': 1, 'font_size': 8})
    no_border = workbook.add_format({'align': 'center', 'border': 0})
    header_cell = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'center', 'border': 1, 'font_size': 7})
    worksheet.set_column('A:A', 11, header_cell)
    worksheet.set_column('B:Y', 3, normal_cell)

    row = 0
    for day in range(0, 7):
        __write_time_headings(worksheet, row, day, header_cell)
        row += 1
        __write_medication_for_day(worksheet, medications, day, row)
        row += len(medications)
        worksheet.set_row(row, None, no_border)
        worksheet.set_row(row+1, None, no_border)
        row += 2

    workbook.close()

def __write_time_headings(worksheet, row, day, format):
    worksheet.write(row, 0, (date.today() + timedelta(days=day)).strftime("%a %d/%m"))
    worksheet.write(row, 1, "12am", format)
    worksheet.write(row, 13, "12pm", format)
    for col in range(1, 12):
        worksheet.write(row, col + 1, str(col) + "am", format)
        worksheet.write(row, col + 13, str(col) + "pm", format)


def __write_medication_for_day(worksheet, medications, day_index, start_row):
    for medication_index, medication in enumerate(medications):
        worksheet.set_row(start_row + medication_index, 23)
        worksheet.write(start_row + medication_index, 0, medication.name)

        time_index = 0
        day = date.today() + timedelta(days=day_index)
        for i, time in enumerate(medication.timetable):
            if time.day == day.day:
                time_index = i
                break

        for hour in range(0, 24):
            if medication.timetable[time_index].hour+1 == hour:
                timestamp = medication.timetable[time_index].strftime("%H:%M") + "\n" if medication.timetable[time_index].minute != 0 else ""
                dose = str(medication.dose) if medication.dose != None else ""
                label = "%s%s %s" % (timestamp, dose, medication.name[0])
                worksheet.write(start_row + medication_index, hour, label)
                time_index += 1

if __name__ == "__main__":
    main()