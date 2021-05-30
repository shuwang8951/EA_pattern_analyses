from EcoCivMdl.code import Utility
import openpyxl
import re


class ChinaDataProcessor:

    def __init__(self):
        # 通用function
        self.ul = Utility.Utility_fc()

        # mode-code pair path
        self.code2mode_path = '../data/distribution/mode-code.xlsx'

        # o_mode 2 code path
        self.data_o_mode2code_path = '../data/distribution/crawled_modes-hand.xlsx'

        # data original mode path
        self.data_o_mode_path = '../data/distribution/ECM_clusterdata_xy.xlsx'

    def process_logic(self):
        # step 1: 填写o-mode 2 mode 信息
        self.write_o_mode2mode()
        # step 2：遍历o-mode 数据，每个o-mode 生成对应mode
        self.write_item_mode()

    def write_item_mode(self):
        # step 1: 载入omode2mode关联dict
        omode2mode_dict = self.generate_omode2mode_dict()

        # step 2: 遍历每个omode，写入mode
        wb = self.ul.load_excel(self.data_o_mode_path)
        sheets = self.ul.load_sheet_list(wb)

        for sheet in sheets:
            omode_list = self.ul.get_col_list_from_sheet(sheet, "H")
            for i in range(len(omode_list)):
                if omode_list[i] in omode2mode_dict.keys():
                        # omode2mode_dict.has_key(str(omode_list[i])):
                    mode = str(omode2mode_dict[omode_list[i]])
                else:
                    mode = ''
                sheet.cell(row=i + 2, column=12, value=mode)
            print(sheet.title + " has been processed!")
        wb.save(self.data_o_mode_path)

    def write_o_mode2mode(self):
        # step 1: 载入code2mode关联信息
        code2mode_dict = self.generate_code2mode_dict()
        # step 2: 循环o-mode2code，匹配每个o-mode对应的mode，生成对应的modelist
        wb = self.ul.load_excel(self.data_o_mode2code_path)

        sheet_omode2code = wb["omode-code"]
        load_code_list = self.ul.get_col_list_from_sheet(sheet_omode2code, "B")
        code_trsf_mode = []
        for i in range(len(load_code_list)):
            code_trsf_mode.append(code2mode_dict[load_code_list[i]])

        # step 3:将对应的modelist写入sheet
        for i in range(len(code_trsf_mode)):
            sheet_omode2code.cell(row=i + 2, column=3, value=str(code_trsf_mode[i]))
        wb.save(self.data_o_mode2code_path)

    def generate_code2mode_dict(self):
        code2mode_dict = {}
        # 载入data
        workbook = self.ul.load_excel(self.code2mode_path)
        sheet_code2mode = workbook["Sheet1"]

        code_list = self.ul.get_col_list_from_sheet(sheet_code2mode, "A")
        mode_list = self.ul.get_col_list_from_sheet(sheet_code2mode, "B")

        for i in range(len(code_list)):
            code2mode_dict[code_list[i]] = mode_list[i]

        # 测试输出
        # for i in code2mode_dict:
        #     print("dict[%s]=" % i, code2mode_dict[i])
        return code2mode_dict

    def generate_omode2mode_dict(self):
        omode2mode_dict = {}
        # 载入data
        workbook = self.ul.load_excel(self.data_o_mode2code_path)
        sheet_omode2mode = workbook["omode-code"]

        omode_list = self.ul.get_col_list_from_sheet(sheet_omode2mode, "A")
        mode_list = self.ul.get_col_list_from_sheet(sheet_omode2mode, "C")

        for i in range(len(omode_list)):
            omode2mode_dict[omode_list[i]] = mode_list[i]

        # 测试输出
        # for i in code2mode_dict:
        #     print("dict[%s]=" % i, code2mode_dict[i])
        return omode2mode_dict


cdp = ChinaDataProcessor()
cdp.process_logic()
