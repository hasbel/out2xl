# Developer: Hassib Belhaj
# Version: V 1.0.2 (28.03.2018)

import os

from outlook import OutlookCalendar
from excel import ExcelSheet
from configuration import Configuration


### LOCATION OF THE CONFIG FILE ###
CONFIG_FILE='out2xl.ini'
## ------------------------------##


def main():
    os.system('mode con: cols=120 lines=300')
    config = Configuration(CONFIG_FILE)
    for calendar in config.calendars:
        sheet = ExcelSheet(calendar[0], calendar[1], config)
        outlook_calendar = OutlookCalendar(config.shared_calendar_owner)
        absence_list = outlook_calendar.get_absence_list(calendar[1], calendar[2])
        for absence in absence_list:
            sheet.mark_absence(absence)


if __name__ == '__main__':
    try:
        main()
    except EnvironmentError as e:
        print("ERROR: {error}".format(error=e))
        print('Execution Stopped !')
    except Exception as e:
        print('Unknown Error Occured: ')
        print(e)
    # Stop the windows command line from closing automatically
    raw_input('Press ENTER to exit')
