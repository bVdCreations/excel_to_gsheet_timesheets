import gspread
from oauth2client.service_account import ServiceAccountCredentials

import json
import datetime


class ToGsheet:

    def __init__(self, year=2018):

        self._scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        self._creden = ServiceAccountCredentials.from_json_keyfile_name('TimeSheetsToDrive.json', self._scope)
        self._client = gspread.authorize(self._creden)
        self._year = year
        self._workfile = self.open_workfile()

    def set_year(self, year:int()):
        """sets the year of the object"""
        self._year = year

    def get_year(self):
        """:return: the year"""
        return self._year

    def get_sheetnames(self):
        """looks for all the sheets names of the current workfile

        :return: a list with all the sheets names
        """
        # return a list with the names of all the sheets in the file
        list_names = list()
        for i in range(len(self._workfile.worksheets())):
            list_names.append(self._workfile.get_worksheet(i).title)

        return list_names

    def open_workfile(self):
        """open the workfile (spreadsheet) by the set year

        :return: the workfile
        """
        # open the worksheet based year that is in the excel file
        try:
            workfile = self._client.open("Timesheets {}".format(self.get_year()))
            return workfile
        except:
            print("create the worksheet")

    def update_timesheets(self,input_timesheets: dict()):

        for title_timesheets, values_timesheet in input_timesheets.items():
            Timesheet(title_timesheets).update_timesheet(values_timesheet)




class Timesheet(ToGsheet):

    def __init__(self, title_timesheet):
        super().__init__()
        self._title_timesheet = title_timesheet.lower()
        self._timesheet = self._open_timesheet()

    def _create_new_timesheet(self):
        # this method creates a new sheet and places the heading for the timesheet
        if self._title_timesheet not in self.get_sheetnames():
            sheet = self._workfile.add_worksheet(self._title_timesheet, 30, 15)
            name_split = self._title_timesheet.split(' ')
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

    def _open_timesheet(self):
        """opens or creates the timesheet by the set title in the object

        :return: returns the opened or created sheet
        """
        if self._title_timesheet in self.get_sheetnames():
            # open sheet
            sheet = self._workfile.worksheet(self._title_timesheet)
        else:
            # create sheet
            sheet = self._create_new_timesheet()
        return sheet

    def update_timesheet(self, time_sheet_input: dict()):
        '''updates every cell of the timesheet with the input dict

        :param time_sheet_input: this is a dict with key as coordinates of a cell
                                   and the value with the corresponding value of the cell
        :return: no return
        '''
        # this puts the readed data from the excel in the rights cell in the sheet in the spreadsheet
        for key_update, value_update in time_sheet_input.items():
            self._open_timesheet().update_acell(key_update, value_update)

    def get_last_entry_row_timesheet(self):
        """looks for the last input in the row
        (staring from row 5 in col A)

        :return:row number of last input
        """
        # find the row number of the last entry in the given column
        row = 5
        last_entry_row = row
        while self._open_timesheet().cell(col=1, row=row).value != '':
            last_entry_row = row
            row += 1
        return last_entry_row

    def get_last_entry_column_timesheet(self, start_column=1):
        '''looks for the last input in the row
                (staring from row 5 in col A)

        :return:row number of last input
        '''
        # find the column number of the last entry the given row
        column = start_column
        last_entry_column = column
        while self._open_timesheet().cell(col=column, row=4).value != '':
            last_entry_column = column
            column += 1
        return last_entry_column

    def title(self):
        """:return title of timesheet"""
        return self._title_timesheet

    def cell(self, row: int(), col: int()):
        """get the cell of the timesheet

        :param row: the row position of the cell
        :param col: the column position of the cell
        :return: the specific object cell
        """
        return self._timesheet.cell(row=row, col=col)


class DaySummary(ToGsheet):

    def __init__(self):
        super().__init__()
        self._day_summary = {'days': {'Monday': 'B', 'Tuesday': 'C', 'Wednesday': 'D', 'Thursday': 'E', 'Friday': 'F',
                                      'Saturday': 'G', 'Sunday': 'H'},
                             'extra_info': {'Total': 'J', 'Extra hours': 'K', 'Average': 'L'}}

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

    def update_summary_week(self, week: "number of the week"):
        """updates the summary formulas in the sheet 'day_summary' in a given week

        :param week: the number of the week
        :param year: the year
        :return:
        """
        sheet_day_summary = self.open_day_summary()
        for key_update, value_update in \
                self.create_summary_formulas(week).items():
            sheet_day_summary.update_acell(key_update, value_update)

    def create_summary_formulas(self, week: "number of the week"):
        """greates a dict with the gsheets formula's for the spefic week

        :param week: the number of the week
        :param year: the year
        :return: dict with the formula's
        """

        summary_formulas = dict()
        timesheet = Timesheet("week {} {}".format(week, self.get_year()))
        timesheet_title = timesheet.title()
        type_of_activity = self.load_type_of_acticity()

        number_week = str(timesheet_title).split(' ')[1]
        column_position = str(int(number_week) + 1)

        # loop true al input value rows in timesheet
        for row in range(5, timesheet.get_last_entry_row_timesheet()+1):
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
                    newform = oldform[:-4] + "+({}-{}) ".format("'"+timesheet_title + "'!D" + str(row),
                                                                "'" + timesheet_title + "'!C" + str(row)) + ")*24"
                    summary_formulas[key_formula] = newform
                else:
                    # create the formula
                    summary_formulas[key_formula] = "=(({}-{}) )*24".format("'"+timesheet_title + "'!D" + str(row),
                                                                            "'" + timesheet_title + "'!C" + str(row))
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
        summary_formulas['A'+column_position] = timesheet_title

        return summary_formulas

    @staticmethod
    def day_of_week(time: '%Y-%m-%d %H:%M:%S'):
        """
        :return: the day of the week from the variable time gives as '%Y-%m-%d %H:%M:%S'
        """
        dayoftheweek = datetime.datetime.strptime(time, '%Y-%m-%d %H:%M:%S').strftime('%A')
        return dayoftheweek

    @staticmethod
    def load_type_of_acticity():
        """:return: dict of types of activity in the company from a json file"""
        with open("type_of_activity.json") as json_data:
            type_of_activity = json.load(json_data)
        return type_of_activity


if __name__ == "__main__":
    pass
