from __future__ import unicode_literals
import datetime


class Absence:
    def __init__(self, appointment_item):
        self.name = self.parse_name(appointment_item.Subject)
        self.start = self.convert_date(appointment_item.Start)
        self.end = self.convert_date(appointment_item.End, is_end=True)
        self.days_list = []
        self.enum_dates()

    def parse_name(self, subject):
        """ For a given absence event subject, return the absentee's name """
        return subject.lower().replace('absence', '').replace(':', '')

    def convert_date(self, date, is_end=False):
        """ Convert a date's string representation to a datetime object

        Is_end is need because outlook save the end date of all day event as 00:00 of the next day. so we need to
        subtract 1 day.
        """
        date = str(date)
        date_obj = datetime.date(2000+int(date[6:8]), int(date[:2]), int(date[3:5]))
        if is_end and date[-8:] == '00:00:00':
            date_obj -= datetime.timedelta(days=1)
        return date_obj

    def enum_dates(self):
        """ Fill up self.days_list with all the dates between self.start and self.end """
        delta = self.end - self.start
        for i in range(delta.days + 1):
            self.days_list.append(self.start + datetime.timedelta(days=i))
