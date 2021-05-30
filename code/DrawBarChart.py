import matplotlib.pyplot as plt

from pyecharts import options as opts
from pyecharts.charts import Bar
import pandas as pd
from pyecharts.globals import CurrentConfig, OnlineHostType
from pyecharts.charts import Grid
from pyecharts.options.global_options import ThemeType
from pyecharts.charts import Pie, Line
import pyecharts
import webbrowser
import time
from pandas import to_datetime
from datetime import datetime


class DrawBarChart:

    def __init__(self):
        # 载入pyecharts文件
        CurrentConfig.ONLINE_HOST = "D:\PYCHARM_WS\EcoCivMdl\pyecharts-assets-master/assets/"

        # 载入数据
        self.data_path = '../data/CEA-DRAW.xlsx'
        data_read = pd.read_excel(self.data_path, sheet_name=0)
        self.CEA_rank = list(data_read['cea-rank'])  # rank
        self.CEA_modes = list(data_read['CEA'])  # mode
        self.CEA_count = list(data_read['count'].values)  # count
        self.CEA_sum = list(data_read['sum'].values)
        self.CEA_4000count = list(data_read['4000count'].values)
        self.CEA_percentage = list(data_read['%'])  # percentage

        # 颜色
        self.colors = ["#5793f3", "#d14a61", "#675bba"]

    def logic(self):
        self.bar_echart()

    def bar_echart(self):
        filepath = r'../data/bar/bar.html'
        bar = Bar(init_opts=opts.InitOpts(width="1600", height="600px"))
        bar.add_xaxis(self.CEA_rank)
        bar.add_yaxis("Record count", y_axis=self.CEA_count, yaxis_index=0, bar_max_width=10)
        bar.extend_axis(
            yaxis=opts.AxisOpts(
                name="Record count",
                type_="number",
                min_=0,
                max_=4000,
                position="left",
                axisline_opts=opts.AxisLineOpts(
                    linestyle_opts=opts.LineStyleOpts(color=self.colors[0])
                )
            )
        )
        bar.set_global_opts(title_opts=opts.TitleOpts(title='', subtitle=''),
                            xaxis_opts=opts.AxisOpts(name_rotate=60, name="Top CEA mode rank",
                                                     axislabel_opts={"rotate": 60}),
                            # yaxis_opts=opts.AxisOpts(
                            #     name="percentage",
                            #     type_="value",
                            #     min_=0,
                            #     max_=100,
                            #     position="right",
                            #     axisline_opts=opts.AxisLineOpts(
                            #         linestyle_opts=opts.LineStyleOpts(color=self.colors[1])
                            #     ),
                            #     splitline_opts=opts.SplitLineOpts(
                            #         is_show=True, linestyle_opts=opts.LineStyleOpts(opacity=1)
                            #     )
                            # ),
                            # tooltip_opts=opts.TooltipOpts(trigger="axis", axis_pointer_type="cross")
                            )

        line = Line()
        line.add_xaxis(self.CEA_rank)
        line.add_yaxis(series_name="percentage", y_axis=self.CEA_percentage,
                       label_opts=opts.LabelOpts(is_show=False))

        bar.overlap(line).render(filepath)

    def bar_matplotlib(self):
        pass


dbc = DrawBarChart()
dbc.logic()
