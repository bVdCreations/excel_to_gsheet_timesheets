import os
import openpyxl


class FindFiles:

    def __init__(self):
        self._patch = "C:\\Users\\bVd\\Desktop\\"
        self._files = dict()
        self._user = "Bastiaan Van Denabeele"

    def get_folder_list(self):
        return os.listdir(self._patch)

    def find_excel(self):
        for file in self.get_folder_list():
            if "Timesheet - {} - Week".format(self._user)in file:
                key = file.split(" - ")[2].split(".")[0]
                self._files.update({key: self._patch+file})
        return self._files


class ReadTimeSheets:

    def __init__(self):
        self._file_list_dict = FindFiles().find_excel()
        self._time_sheet_input = dict()

    def get_files_dir(self):
        return self._file_list_dict

    def get_sheets(self):
        returndict = dict()
        for keys, value in self._file_list_dict.items():
            returndict.update({keys: openpyxl.load_workbook(value).get_sheet_by_name('Timesheet')})
        return returndict

    def get_sheet_input(self):
        sheet_inputs = dict()
        for key, value in self.get_sheets().items():
            sheet_inputs.update({key: dict()})
            for rowOfCellObjects in value['A5':self.get_last_entry_row(value)]:
                for cellObj in rowOfCellObjects:
                    if cellObj.value is not None:
                        sheet_inputs.get(key).update({cellObj.coordinate: cellObj.value})

        return sheet_inputs

    def get_last_entry_row(self, sheet_object: openpyxl):
        for row in sheet_object.columns:
            for cell in row:
                if cell.value is None and cell.coordinate != 'A3':
                    return 'L{}'.format(int(cell.coordinate.strip('A'))-1)

    def get_sheet_input2(self):
        sheet_inputs = dict()
        for key, value in self.get_sheets().items():
            sheet_inputs.update({key: dict()})
            for rowOfCellObjects in value['A1':'L4']:
                for cellObj in rowOfCellObjects:
                    if cellObj.value is not None:
                        sheet_inputs.get(key).update({cellObj.coordinate: cellObj.value})

        return sheet_inputs


if __name__ == "__main__":
    rt = ReadTimeSheets()
    rt5 = ReadTimeSheets()
    print(rt.get_sheet_input2())