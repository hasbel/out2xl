import ConfigParser


class Configuration:
    def __init__(self, config_file):
        try:
            self.config = ConfigParser.RawConfigParser()
            self.config.read(config_file)
            self.shared_calendar_owner = self.config.get('GENERAL', 'SHARED_CALENDAR_OWNER')
            self.first_name_row = self.config.getint('GENERAL', 'FIRST_NAMES_ROW')
            self.last_name_row = self.config.getint('GENERAL', 'LAST_NAMES_ROW')
            self.first_calendar_column = self.config.get('GENERAL', 'FIRST_CALENDAR_COLUMN')
            self.calendars =[]
            self.process_calendars()
        except Exception as e:
            print(e)
            raise EnvironmentError('Could not correctly parse the configuration file. '
                                   'Is in the same folder as this program? Does it have the correct format?')

    def process_calendars(self):
        """ Fill up self.calendars

        For each defined calendar in the config file, create a tuple with the format:
        (calendar_name, calendar_start_date, Calendar_end_date)
        Add this tuple to the self.calendars list.
        """
        for section in self.config.sections()[1:]:
            begin = self.config.get(section, 'FIRST_DATE')
            end = self.config.get(section, 'LAST_DATE')
            self.calendars.append((section, begin, end))