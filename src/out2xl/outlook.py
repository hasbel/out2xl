from __future__ import unicode_literals
import win32com.client

from absence import Absence


class OutlookCalendar:
    def __init__(self, owner):
        self.app = win32com.client.Dispatch("Outlook.Application")
        self.namespace = self.app.GetNamespace("MAPI")
        self.recipient = self.namespace.createRecipient(owner)
        self.calendar = self.load(self.recipient)

    def load(self, owner):
        if not owner.Resolve():  # Resolve needs to be run, and needs to succeed.
            print ("ERROR: Could not resolve calendar creator name. Quitting...")
            exit()
        return self.namespace.GetSharedDefaultFolder(owner, 9)  # returns Folder object

    def get_absence_list(self, start, end):
        appointments = self.calendar.Items.Restrict(
            "[START] >= '{start_date}' AND [END] <= '{end_date}'".format(start_date=start, end_date=end))
        absence_list = []
        for item in appointments:
            if "absence" in item.Subject.lower():
                absence = Absence(item)
                absence_list.append(absence)
        return absence_list
