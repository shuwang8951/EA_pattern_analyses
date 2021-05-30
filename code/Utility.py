import openpyxl


class Utility_fc:
    def __init__(self):
        pass

    def load_dict_file(self, path):
        mode_re_list = []
        try:
            file = open(path, 'r', encoding='utf8')

            for pattern in file.readlines():
                mode_re_list.append(pattern.strip())
            mode_re_list.pop(0)
        except:
            print("load dict file error!")
        return mode_re_list

    def load_excel(self, file_path):
        try:
            wb = openpyxl.load_workbook(file_path)
            print("data file read successfully.")
            return wb
        except Exception as res:
            print("data file read unsuccessfully.")
            print(res)

    def load_sheet_list(self, wb):
        sheets = []
        if wb:
            shs = wb.get_sheet_names()
            for i in range(len(shs)):
                sheets.append(wb.get_sheet_by_name(shs[i]))
        return sheets

    def get_list_from_sheet_cell(self, sheet_cell):
        list = []
        if sheet_cell == '[]':
            return list
        else:
            # 去括号
            temp = sheet_cell[2:len(sheet_cell) - 2]

            if "', '" in temp:
                list = temp.split("', '")
                return list
            else:
                list.append(temp)
                return list

    def get_col_list_from_sheet(self, sheet, col_code):
        col = sheet[col_code]

        datalist = []

        for x in range(len(col)):
            if x > 0:
                value = col[x].value
                if value is None:
                    pass
                else:
                    value = str(value)
                    datalist.append(value)

        return datalist