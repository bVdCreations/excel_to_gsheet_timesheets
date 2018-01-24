from Send_Timesheet_Gsheet import ToGsheet
from Read_Timesheet import ReadTimeSheets


if __name__ == "__main__":
    sorted_year = dict()
    for key, value in ReadTimeSheets().get_sheet_input().items():
        if key.split(" ")[2] not in sorted_year.keys():
            sorted_year[key.split(" ")[2]] = {key:value}
        else:
            sorted_year[key.split(" ")[2]].update({key:value})
        print("key: {}::".format(key))
        print("     value: {}".format(value))
        #ToGsheet(key, value).update_summary_day()


    for year,timesheets_input in sorted_year.items():
        ToGsheet(year).update_timesheets(timesheets_input)