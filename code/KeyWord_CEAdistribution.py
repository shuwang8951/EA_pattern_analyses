from EcoCivMdl.code import Utility
import jieba.analyse as ana


class SupplementKeyWord:

    def __init__(self):
        # 通用function
        self.ul = Utility.Utility_fc()

        # data path
        self.data_path = '../data/CEA_Distribution_2018_2020.xlsx'

    def process_logic(self):
        # 载入数据
        wb = self.ul.load_excel(self.data_path)
        ws = wb["数据集"]

        # 获取MODE列
        MODE_list = self.get_col_list_from_sheet(ws, "I")

        for i in range(len(MODE_list)):
            print(str(i + 1) + " is processing......")  # 进度控制

            keywords = []
            keywords_str = ""

            # mode作为第一个关键词
            mode = MODE_list[i]
            keywords.append(mode)

            # 获取content
            content = ws["F"][i + 1].value

            # jieba 利用Textrank解析关键词
            keyword_jieba = ana.textrank(content, topK=4)

            for kw in keyword_jieba:
                keywords.append(kw)

            # 构建keyword str
            for kw_e in keywords:
                keywords_str = keywords_str + kw_e + "；"

            # 写入keyword 列
            ws.cell(row=i + 1, column=15, value=keywords_str)

        # 保存
        wb.save(self.data_path)

    def get_col_list_from_sheet(self, sheet, col_code):
        col = sheet[col_code]
        datalist = []
        for x in range(len(col)):
            if x > 0:
                value = col[x].value
                if value is None:
                    datalist.append("")
                else:
                    value = str(value)
                    datalist.append(value)
        return datalist


skw = SupplementKeyWord()
skw.process_logic()
