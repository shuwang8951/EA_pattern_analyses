from EcoCivMdl.code import CrawlRbt
from EcoCivMdl.code import MinistryWebdata
from EcoCivMdl.code import MinistryDataParse
from EcoCivMdl.code import YangshiNewsdata


class TextCrawler:
    def __init__(self):
        pass

    def crawl(self):
        # ministry article crawl
        # mw = MinistryWebdata.MinistryWebdata()
        # mw.crawl_ministry_web()
        # mw.crawl_each_article()
        # mw.crawl_title_and_date()

        # mdp = MinistryDataParse.MinistryDataParse()

        # yangshi news data crawl
        ysnews = YangshiNewsdata.YangshiNewsdata()
        ysnews.crawl_yangshinews_web()
