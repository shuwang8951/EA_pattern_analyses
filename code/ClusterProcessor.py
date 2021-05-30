from EcoCivMdl.code import Utility
import re


class ClusterProcessor:

    def __init__(self):
        # 通用function
        self.ul = Utility.Utility_fc()

        # data path
        self.data_path = '../data/ECM_data_all-original.xlsx'

        # written data path
        self.written_data_path = '../data/ECM_clusterdata.xlsx'

        # hand excel data path
        self.crawled_modes_data_path = '../data/crawled_modes.xlsx'

        # filter list
        self.filter_list_path = '../data/dict_filter.txt'
        self.filter_list = self.ul.load_dict_file(self.filter_list_path)

    def generate_hand_excel(self):
        # step 1: 遍历爬取数据
        workbook = self.ul.load_excel(self.written_data_path)
        sheets = self.ul.load_sheet_list(workbook)
        # 载入待写入数据
        wb_m = self.ul.load_excel(self.crawled_modes_data_path)

        # 存储含重复的mode，用于计算mode频率
        modes_list = []
        # 存储 sentence of the mode
        modes_sentence = {}

        # 遍历每个sheet
        for sheet in sheets:
            # 遍历每条记录
            for index in range(2, sheet.max_row + 1):
                each_mode = str(sheet.cell(row=index, column=8).value)
                sentence = str(sheet.cell(row=index, column=7).value).strip()

                if each_mode not in modes_list:
                    # 存储sentence of the mode
                    modes_sentence[each_mode] = sentence

                # 存到modes list中
                modes_list.append(each_mode)
                print(sheet.title + " : " + str(index) + "/" + str(sheet.max_row + 1))

        # 新建待写入sheet
        w_sheet = wb_m.create_sheet("mode statistics", 0)
        index = 1
        for mode in modes_sentence:
            print('正在存储mode：'+str(index))
            # mode
            w_sheet.cell(row=index + 1, column=1, value=str(mode))
            # fre
            w_sheet.cell(row=index + 1, column=2, value=str(modes_list.count(mode)))
            # sentence
            w_sheet.cell(row=index + 1, column=3, value=str(modes_sentence[mode]))
            # counter
            index = index + 1
        wb_m.save(self.crawled_modes_data_path)

    def process_filter_data(self):
        # step 1：载入处理数据
        workbook = self.ul.load_excel(self.data_path)
        sheet_list = self.ul.load_sheet_list(workbook)
        # 载入待写数据
        w_wb = self.ul.load_excel(self.written_data_path)

        # step 2: 每条记录过滤
        for sheet in sheet_list:
            # 新建待写入sheet
            w_sheet = w_wb.create_sheet(sheet.title, 0)
            written_row_index = 2
            for index in range(2, sheet.max_row + 1):
                modes_in_sen = str(sheet.cell(row=index, column=10).value)
                parse_list = self.ul.get_list_from_sheet_cell(modes_in_sen)
                index_mode_set = set()

                # 拆开每个模式
                if parse_list:
                    for each_mode_in_sen in parse_list:
                        each_mode = each_mode_in_sen.split('@@')[0]  # 获取模式
                        mode_sen = each_mode_in_sen.split('@@')[1]  # 获取所在短句子

                        # 单条新闻中，仅保留一个模式
                        if each_mode not in index_mode_set:
                            index_mode_set.add(each_mode)
                            # 获取所在长句子
                            content = sheet.cell(row=index, column=7).value
                            long_sen = self.get_long_sentence(each_mode, content)

                            # 模式，删除掉过滤词汇
                            if self.contained_by_filter_word_list(each_mode) is False:
                                # 写入每个模式关联的信息
                                # sheet_id
                                sheet_id = sheet.cell(row=index, column=1).value
                                w_sheet.cell(row=written_row_index, column=1, value=str(sheet_id))
                                # url
                                url = sheet.cell(row=index, column=2).value
                                w_sheet.cell(row=written_row_index, column=2, value=str(url))
                                # title
                                title = sheet.cell(row=index, column=3).value
                                w_sheet.cell(row=written_row_index, column=3, value=str(title))
                                # date
                                date = sheet.cell(row=index, column=4).value
                                w_sheet.cell(row=written_row_index, column=4, value=str(date))
                                # content
                                w_sheet.cell(row=written_row_index, column=5, value=str(content))
                                # short_sentence
                                w_sheet.cell(row=written_row_index, column=6, value=str(mode_sen))
                                # long_sentence
                                w_sheet.cell(row=written_row_index, column=7, value=str(long_sen))
                                # mode_name
                                w_sheet.cell(row=written_row_index, column=8, value=str(each_mode))
                                # place name
                                place_name = sheet.cell(row=index, column=17).value
                                w_sheet.cell(row=written_row_index, column=9, value=str(place_name))

                                # written row index
                                written_row_index = written_row_index + 1
            w_wb.save(self.written_data_path)

            print(sheet.title + " has been processed!")

    def contained_by_filter_word_list(self, mode):
        contained = False
        for fw in self.filter_list:
            if fw == mode:
                contained = True
                break
        return contained

    def get_long_sentence(self, mode, content):
        long_sentence = ''
        content = re.sub('[。”]{2}', '#', content)
        content = re.sub('[？”]{2}', '{', content)
        content = re.sub('[！”]{2}', '}', content)
        sentences = re.split('(#|{|}|。|！|？|!|\?)', content)
        for sent in sentences:
            if mode in sent:
                long_sentence = sent
                break
        return long_sentence


cp = ClusterProcessor()
# cp.process_filter_data()
cp.generate_hand_excel()
