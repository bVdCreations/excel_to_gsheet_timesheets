from Send_Timesheet_Gsheet import TimeSheetToGsheet
from Read_Timesheet import ReadTimeSheets


if __name__ == "__main__":
    for key, value in ReadTimeSheets().get_sheet_input().items():
        TimeSheetToGsheet(key, value).update_summary_day()
