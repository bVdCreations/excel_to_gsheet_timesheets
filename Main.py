from Send_Timesheet_Gsheet import ToGsheet
from Read_Timesheet import ReadTimeSheets


if __name__ == "__main__":

    for year, timesheets_input in ReadTimeSheets().get_sheet_input().items():
        ToGsheet(year).update_timesheets(timesheets_input)
