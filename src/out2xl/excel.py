from __future__ import unicode_literals
import win32com.client
import datetime


class ExcelSheet:
    def __init__(self, sheet_name, start_date, config):
        self.app = win32com.client.Dispatch('Excel.Application')
        self.workbook = self.app.ActiveWorkbook
        self.sheet_name = sheet_name
        self.sheet = self.load()
        self.start_date = start_date
        self.first_name_row = config.first_name_row
        self.last_name_row = config.last_name_row
        self.first_column = convert_column_to_int(config.first_calendar_column)
        self.names = {}
        self.read_names()

    def load(self):
        """ Load the Excel sheet that corresponds the self.sheet_name """
        try:
            return self.app.ActiveWorkbook.Sheets.Item(self.sheet_name)
        except:
            del self.app
            raise EnvironmentError("Excel is either not running, or no Excel sheet called {sheet} Could not be found."
                                   .format(sheet=self.sheet_name.decode('iso-8859-1').encode(errors='replace')))

    def read_names(self):
        """ Fill up the self.names dictionary with all the names available in the Excel sheet and their Row number """
        for row in range(self.first_name_row, self.last_name_row+1):
            name_cell = self.sheet.Cells(row, 'D')
            if name_cell.Value:
                self.names[name_cell.Value.lower()] = row

    def find_row(self, name):
        """ For a given name, return the corresponding row number if it exists, otherwise return -1 """
        for employee_name in self.names:
            employee_name_parts = employee_name.replace(',', ' ').split()
            name_match = True
            for name_part in employee_name_parts:
                if name_part not in name:
                    name_match = False
            if name_match:
                return self.names[employee_name]
        return -1

    def find_column(self, date):
        """ Return the Excel column number corresponding to the given date """
        start_year = int(self.start_date.split('.')[2])
        start_month = int(self.start_date.split('.')[1])
        start_day = int(self.start_date.split('.')[0])
        excel_calendar_first_date = datetime.date(start_year, start_month, start_day)
        time_delta = date - excel_calendar_first_date
        return time_delta.days + self.first_column

    def mark_absence(self, absence):
        """ Mark the given absence in the Excel sheet """
        row = self.find_row(absence.name)
        if row < 0:  # Name was not found in the Excel sheet
            print("no match with name: {name}".format(name=absence.name.encode(errors='replace')))
            print("Please make sure it exists in the Excel sheet {sheet}."
                  .format(sheet=self.sheet_name.decode('iso-8859-1').encode(errors='replace')))
            print("Absence: {name} - {begin} to {end} will be skipped."
                  .format(name=absence.name.encode(errors='replace'), begin=absence.start, end=absence.end))
            print('------------------------')
            return
        for day in absence.days_list:
            column = self.find_column(day)
            if column < self.first_column:
                return
            cell = self.sheet.Cells(row, column)
            cell.Value = 0.0
            # Do not recolor violet cells. Colors: blue=5 , violet=39 , white=-4142
            if cell.Interior.ColorIndex != 39:
                cell.Interior.ColorIndex = 5

def convert_column_to_int(column):
    """ For a given Excel column name, return the corresponding column number (Ex: 'AA' -> 27) """
    column_number = 0
    multiplier = 1
    for character in column:
        column_number += (ord(character.lower()) - 96) * multiplier
        multiplier *= 26
    return  column_number
