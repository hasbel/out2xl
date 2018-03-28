from __future__ import unicode_literals
import win32com.client

from absence import Absence


class OutlookCalendar:
    def __init__(self, owner):
        self.app = win32com.client.Dispatch('Outlook.Application')
        self.calendar = self.load(owner)

    def load(self, owner):
        """ Return the Outlook COM calendar object belonging to the given calendar owner"""
        try:
            namespace = self.app.GetNamespace('MAPI')
            recipient = namespace.createRecipient(owner)
            if not recipient.Resolve():  # Resolve needs to be run, and needs to succeed.
                raise EnvironmentError('Could not resolve calendar creator Email. Is the correct Email provided ?')
            return namespace.GetSharedDefaultFolder(recipient, 9)  # returns Folder object
        except AttributeError:
            del self.app
            raise EnvironmentError('Could not connect to Outlook. Is it running ?')

    def get_absence_list(self, start, end):
        """ Get a list of Absence objects that represent all absences in the calendar

        To count as an absence, the event must be between the start and end date, and it's subject
        must start with the world 'Absence'
        """
        appointments = self.calendar.Items.Restrict(
            "[START] >= '{start_date}' AND [END] <= '{end_date}'".format(start_date=start, end_date=end))
        absence_list = []
        for item in appointments:
            if 'absence' in item.Subject.lower():
                absence = Absence(item)
                absence_list.append(absence)
        return absence_list
