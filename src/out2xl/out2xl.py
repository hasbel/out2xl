# Developer: Hassib Belhaj
# Version: V 0.1.4 (12.01.2018)

from __future__ import unicode_literals

from outlook import OutlookCalendar
from excel import ExcelSheet


# CONFIGURATION:
SHEET_NAME = "1. Halbjahr 2018"
CALENDAR_END = "2018-06-30"
CALENDAR_START = "2018-01-01"
SHARED_CALENDAR_OWNER = "family-name, name"
################


def main():
    sheet = ExcelSheet(SHEET_NAME)
    calendar = OutlookCalendar(SHARED_CALENDAR_OWNER)
    absence_list = calendar.get_absence_list(CALENDAR_START, CALENDAR_END)
    for absence in absence_list:
        sheet.mark_absence(absence)
    raw_input('Press ENTER to exit')


if __name__ == '__main__':
    main()
