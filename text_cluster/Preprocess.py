from EcoCivMdl.code import Utility
import jieba
from sklearn.feature_extraction.text import TfidfVectorizer
import numpy as np
from sklearn.metrics.pairwise import cosine_similarity

from scipy.cluster.hierarchy import ward, dendrogram, linkage
import matplotlib.pyplot as plt


class PreProcess:
    def __init__(self):
        # 通用function
        self.ul = Utility.Utility_fc()

        # clusterdata path
        self.cluster_data_path = '../data/clusterdata-pre.xlsx'

        # clusterdata path
        self.cluster_txt_path = '../data/clusterdata.txt'

        # cluster segment corpus path
        self.cluster_segment_txt_path = '../data/cluster_seg_data.txt'

        # stop word list path
        self.stop_words_txt_path = '../data/中文常用过滤词表.txt'

    def pre_processes(self):
        # step 1: excel to txt
        # self.trans2input()
        # step 2: txt 分词，去除停用词，形成corpus txt格式
        # self.write_segment_txt(self.cluster_txt_path, self.cluster_segment_txt_path)
        # step 3:生成tf-idf权重下的VSM词袋空间
        self.generate_tf_idf_vector()

    def trans2input(self):
        """
        pre step 1
        从excel中，转换成input格式txt
        :return:无
        """
        # load data
        data_workbook = self.ul.load_excel(self.cluster_data_path)
        data_sheet = data_workbook["Sheet2"]

        # document list
        document_list = []

        # 遍历每条记录
        for index in range(2, data_sheet.max_row + 1):
            # 获取文本
            each_sentence = str(data_sheet.cell(row=index, column=7).value)
            # 去除换行
            each_document = each_sentence.replace('\n', '').replace('\r', '')
            document_list.append(each_document)

        # 把document list写入txt
        self.write_txt_from_list(self.cluster_txt_path, document_list)

    def write_txt_from_list(self, txt_path, input_list):
        """
        根据读取的document list，和txt路径，写入document语料库
        :param txt_path: document语料库路径
        :param input_list: 读入的document list
        :return:无
        """
        f = open(txt_path, "w", encoding='utf-8')
        if input_list:
            for input in input_list:
                f.write(str(input) + "\n")
            f.close()

    def write_segment_txt(self, txt_path, segment_txt_path):
        with open(txt_path, "r", encoding='utf-8') as f_input:
            index = 1
            for input_each_document in f_input:
                # 用结巴 全模式，分词
                seg_list = jieba.cut(input_each_document, cut_all=True)
                document_seg = " ".join(seg_list)
                # # 转换为list
                # jieba_seg_list = list(seg_list)
                # # 存储document 字符串
                # seg_string = ''
                # # 删除stop words
                # for seg_word in jieba_seg_list:
                #     if seg_word not in stop_word_list:
                #         seg_string = seg_string + seg_word + ' '

                print(str(index))
                # 写入文件
                open(segment_txt_path, "a", encoding='utf-8').write(document_seg.strip() + "\n")
                index = index + 1

        f_input.close()

    def load_txt_titles(self):
        # load data
        data_workbook = self.ul.load_excel(self.cluster_data_path)
        data_sheet = data_workbook["Sheet2"]
        titles = self.ul.get_col_list_from_sheet(data_sheet, "H")
        return titles

    def generate_tf_idf_vector(self):
        corpus = []
        # 生成corpus
        with open(self.cluster_segment_txt_path, 'r', encoding='utf-8') as f:
            for line in f:
                corpus.append(line)

        # 逐行读取生成的document corpus
        stop_word_list = self.ul.load_dict_file(self.stop_words_txt_path)

        vectorizer = TfidfVectorizer(stop_words=stop_word_list, max_features=3)
        tfidf = vectorizer.fit_transform(corpus)
        print(vectorizer.get_feature_names())
        print(tfidf.shape)
        # tfidf_array = tfidf.toarray()
        # U, s, A = np.linalg.svd(tfidf_array)
        # print(s)

        # 计算文档相似性
        dist = 1 - cosine_similarity(tfidf)

        # 获得分类

        plt.rcParams['font.sans-serif'] = ['Microsoft YaHei']  # 用来正常显示中文标签
        # Perform Ward's linkage on a condensed distance matrix.
        # linkage_matrix = ward(dist) #define the linkage_matrix using ward clustering pre-computed distances
        # Method 'ward' requires the distance metric to be Euclidean
        linkage_matrix = linkage(dist, method='ward', metric='euclidean', optimal_ordering=False)
        # Z[i] will tell us which clusters were merged, let's take a look at the first two points that were merged
        # We can see that ach row of the resulting array has the format [idx1, idx2, dist, sample_count]

        # 输出分类名称
        # print(linkage_matrix)
        # for index, title in enumerate(titles):
        #   print(index, title)

        # 可视化
        titles = self.load_txt_titles()
        plt.figure(figsize=(25, 10))
        plt.title('中文文本层次聚类树状图')
        plt.xlabel('微博标题')
        plt.ylabel('距离（越低表示文本越类似）')
        dendrogram(
            linkage_matrix,
            labels=titles,
            leaf_rotation=-70,  # rotates the x axis labels
            leaf_font_size=12  # font size for the x axis labels
        )
        plt.show()
        plt.close()
