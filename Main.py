from Send_Timesheet_Gsheet import ToGsheet
from Read_Timesheet import ReadTimeSheets
from Send_Timesheet_Gsheet import DaySummary


if __name__ == "__main__":
    
    print('started')
    for year, timesheets_input in ReadTimeSheets().get_sheet_input().items():
        
        ToGsheet(year).update_timesheets(timesheets_input)
        print('updated the sheet {} to gheet {}'.format(timesheets_input.keys() ,year))
        for key in timesheets_input.keys():
            try:
                number = int(key.split(' ')[1])
                DaySummary().update_summary_week(number)
                print('update summary week for week {}'.format(number))
            except ValueError:
                raise ValueError('can not confert {} to a number'.format(key))


