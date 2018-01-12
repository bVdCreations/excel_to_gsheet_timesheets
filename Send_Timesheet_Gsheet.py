import gspread
from oauth2client.service_account import ServiceAccountCredentials


class TimeSheetToGsheet:

    def __init__(self, name: str(), update_input: dict()):

        self._scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        self._creden = ServiceAccountCredentials.from_json_keyfile_name('TimeSheetsToDrive.json', self._scope)
        self._client = gspread.authorize(self._creden)
        self._time_sheet_input = update_input
        self._name = name
        self._workfile = self.open_spreadsheet()
        self.update_timesheet()

    def get_sheetnames(self):
        # return a list with the names of all the sheets in the file
        list_names = list()
        for i in range(len(self._workfile.worksheets())):
            list_names.append(self._workfile.get_worksheet(i).title)

        return list_names

    def open_spreadsheet(self):
        # open the worksheet based year that is in the excel file
        try:
            spreadsheet = self._client.open("Timesheets {}".format(self._name.split(' ')[2]))
            return spreadsheet
        except:
            print("create the worksheet")

    def create_new_timesheet(self):
        # this method creates a new sheet and places the heading for the timesheet
        if self._name not in self.get_sheetnames():
            sheet = self._workfile.add_worksheet(self._name, 30, 15)
            name_split = self._name.split(' ')
            cells = {'A1': 'Employee', 'B1': 'Bastiaan Van Denabeele', 'C1': 'Initials', 'D1': 'BVD', 'E1': 'Number',
                     'F1': '51', 'A2': 'Year', 'B2': name_split[2], 'C2': 'Week', 'D2': name_split[1],
                     'A4': 'Type of activity', 'B4': 'Date', 'C4': 'From', 'D4': 'Until', 'E4': 'Project',
                     'F4': 'Transport', 'G4': 'Travel from', 'H4': 'Travel to',
                     'I4': 'Location', 'J4': 'Comments internal', 'K4': 'Comments visible for customer'}

            for key_cell, value_cell in cells.items():
                sheet.update_acell(key_cell, value_cell)

            return sheet
        else:
            raise AttributeError("trying to create a sheet that already exists")


    def open_timesheet(self):

        if self._name in self.get_sheetnames():
            # open sheet
            sheet = self._workfile.worksheet(self._name)
        else:
            # create sheet
            sheet = self.create_new_timesheet()

        return sheet

    def update_timesheet(self):
        # this puts the readed data from the excel in the rights cell in the sheet in the spreadsheet
        for key_update, value_update in self._time_sheet_input.items():
            self.open_timesheet().update_acell(key_update, value_update)


if __name__ == "__main__":
    pass