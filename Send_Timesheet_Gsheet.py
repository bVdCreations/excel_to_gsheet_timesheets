import gspread
from oauth2client.service_account import ServiceAccountCredentials

import json
import datetime


class TimeSheetToGsheet:

    def __init__(self, name: str(), update_input: dict()):

        self._scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        self._creden = ServiceAccountCredentials.from_json_keyfile_name('TimeSheetsToDrive.json', self._scope)
        self._client = gspread.authorize(self._creden)
        self._time_sheet_input = update_input
        self._name = name.lower()
        self._workfile = self.open_spreadsheet()
        self.update_timesheet()
        self._day_summary = {'days': {'Monday': 'B', 'Tuesday': 'C', 'Wednesday': 'D', 'Thursday': 'E', 'Friday': 'F',
                                      'Saturday': 'G', 'Sunday': 'H'},
                             'extra_info': {'Total': 'J', 'Extra hours': 'K', 'Average': 'L'}}

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

    def open_timesheet(self, ):
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

    def get_last_entry_row_timesheet(self):
        # find the row number of the last entry in the given column
        row = 5
        last_entry_row = row
        while self.open_timesheet().cell(col=1, row=row).value != '':
            last_entry_row = row
            row += 1
        return last_entry_row

    def get_last_entry_column_timesheet(self, start_column=1):
        # find the column number of the last entry the given row
        column = start_column
        last_entry_column = column
        while self.open_timesheet().cell(col=column, row=4).value != '':
            last_entry_column = column
            column += 1
        return last_entry_column

    def open_day_summary(self):

        if "Day_Summary" in self.get_sheetnames():
            # open sheet
            sheet = self._workfile.worksheet("Day_Summary")
        else:
            # create sheet
            sheet = self.create_new_day_summary()

        return sheet

    def create_new_day_summary(self):
        # this method creates a new sheet and places the heading for the day summary sheet
        if "Day_Summary" not in self.get_sheetnames():
            sheet = self._workfile.add_worksheet("Day_Summary", 60, 15)

            for item in self._day_summary.values():
                for key_cell, value_cell in item.items():
                    sheet.update_acell((value_cell+'1'), key_cell)

            return sheet
        else:
            raise AttributeError("trying to create a sheet that already exists")

    def update_summary_day(self):
        sheet_day_summary = self.open_day_summary()
        for key_update, value_update in self.create_summary_formulas().items():
            sheet_day_summary.update_acell(key_update, value_update)

    def create_summary_formulas(self):

        summary_formulas = dict()

        type_of_activity = self.load_type_of_acticity()
        timesheet = self.open_timesheet()

        number_week = self._name.split(' ')[1]
        column_position = str(int(number_week) + 1)

        # loop true al input value rows in timesheet
        for row in range(5, self.get_last_entry_row_timesheet()+1):
            # define the cell coordinates in the sheet day summary
            day = self.day_of_week(timesheet.cell(row=row, col=2).value)
            key_formula = self._day_summary.get("days").get(day) + column_position
            # check if the activity (located in the A column) is a working activitiy
            if type_of_activity.get(timesheet.cell(row=row, col=1).value) in ["Working",
                                                                              "SLA Fee", "Training"]:
                # check if the key is already in dict
                if key_formula in summary_formulas.keys():
                    # update the formula
                    oldform = summary_formulas[key_formula]
                    newform = oldform[:-4] + "+({}-{}) ".format("'"+self._name + "'!D" + str(row),
                                                                "'" + self._name + "'!C" + str(row)) + ")*24"
                    summary_formulas[key_formula] = newform
                else:
                    # create the formula
                    summary_formulas[key_formula] = "=(({}-{}) )*24".format("'"+self._name + "'!D" + str(row),
                                                                            "'" + self._name + "'!C" + str(row))
            elif type_of_activity.get(timesheet.cell(row=row, col=1).value) == "Day Off":
                if timesheet.cell(row=row, col=1).value == "Day off (special reason, describe in comments)":
                    summary_formulas[key_formula] = "read the comment for more info"
                else:
                    summary_formulas[key_formula] = 8

        # add the extra
        summary_formulas[self._day_summary.get('extra_info').get('Extra hours') + column_position] = \
            "=J{}-8*{}".format(column_position, len(summary_formulas))
        summary_formulas[self._day_summary.get('extra_info').get('Total') + column_position] =\
            "=SUM(B{}:H{})".format(column_position, column_position)
        summary_formulas[self._day_summary.get('extra_info').get('Average') + column_position] = \
            "=AVERAGE(B{}:H{})".format(column_position, column_position)
        summary_formulas['A'+column_position] = self._name

        return summary_formulas

    @staticmethod
    def day_of_week(time):
        dayoftheweek = datetime.datetime.strptime(time, '%Y-%m-%d %H:%M:%S').strftime('%A')
        return dayoftheweek

    @staticmethod
    def load_type_of_acticity():
        with open("type_of_activity.json") as json_data:
            type_of_activity = json.load(json_data)
        return type_of_activity


if __name__ == "__main__":
    pass
