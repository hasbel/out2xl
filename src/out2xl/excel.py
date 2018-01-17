from __future__ import unicode_literals
import win32com.client
import datetime


class ExcelSheet:
    def __init__(self, sheet_name):
        self.app = win32com.client.Dispatch("Excel.Application")
        self.sheet = self.load(sheet_name)
        self.names = {}
        for row in range(8, 53):
            cell = self.sheet.Cells(row, "D")
            self.names[cell.Value.lower()] = row

    def load(self, name):
        try:
            return self.app.ActiveWorkbook.Sheets.Item(name)
        except AttributeError:
            print("Excel is either not running, or a sheet called '2. Halbjahr 2017' Could not be found.")
            exit()

    def find_row(self, name):
        # divide name into words (first name, second name, surname , etc. )
        name = name.replace(",", " ")
        name = name.split()

        possible_matches = {}
        # iterate over all the employee names found in the excel sheet calendar
        for employee_name in self.names:
            # if any part of the name matches, add it. Necessary as name are not always added correctly
            for name_part in name:
                if name_part in employee_name:
                    if employee_name in possible_matches:
                        possible_matches[employee_name] += 1
                    else:
                        possible_matches[employee_name] = 1

        if len(possible_matches) > 1:
            employee = max(possible_matches, key=possible_matches.get)
            print ("Warning, possible conflict. returning user {match} for name {name}.".format(match=employee, name=str(name)))

        elif len(possible_matches) == 0:
            print ("no match with name: {name}".format(name=str(name)))
            print ("This name will be skipped. Please make sure it exists in the Excel sheet.")
            return -1

        else:
            employee = possible_matches.keys()[0]

        return self.names[employee]

    def find_column(self, date):
        # TODO: GET this value from CALENDAR_START
        excel_calendar_first_date = datetime.date(2018, 1, 1)
        time_delta = date - excel_calendar_first_date
        return time_delta.days + 6  # The first calendar day is Column F. (F = 6)

    def mark_absence(self, absence):
        row = self.find_row(absence.name)
        if row == -1:  # Name was not found in the Excel sheet
            return
        for day in absence.days_list:
            column = self.find_column(day)
            cell = self.sheet.Cells(row, column)
            cell.Value = 0.0
            cell.Interior.ColorIndex = 5  # blue = 5 , violet = 39 , white = -4142