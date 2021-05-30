from goose3 import Goose
from goose3.text import StopWordsChinese
import requests
from fake_useragent import UserAgent
import gc


class CrawlRbt:
    def __init__(self):
        self.ua = UserAgent()

        self.g = Goose({'browser_user_agent': 'Mozilla', 'stopwords_class': StopWordsChinese})

    def get_text(self, url):
        headers = {
            "User-Agent": self.ua.random,
            'Connection': 'close',
        }
        content = ''
        try:
            res = requests.get(url, headers=headers, timeout=8)
        except:
            content = url
            print('未获取链接！')
        if content != url:
            print(res.text)


    def get_text_bygoose(self, url, re_crawlnumber):
        headers = {
            "User-Agent": self.ua.random,
            'Connection': 'close',
        }
        try:
            res = requests.get(url, headers=headers, timeout=8)
            print('提取正文...from：' + url)
            print(res.status_code)
            if res.status_code == 200:
                article = self.g.extract(url)
                content = article.cleaned_text
                if content == None and re_crawlnumber > 0:
                    # self.get_text_bygoose(url, re_crawlnumber - 1)
                    content = url
                    print('该链接内容读取超时！')


            else:
                content = url
                if re_crawlnumber > 0:
                    # self.get_text_bygoose(url, re_crawlnumber - 1)
                    content = url
                    print('该链接内容读取超时！')
        except:
            content = url
            print('该链接内容读取超时！')
            if re_crawlnumber > 0:
                # self.get_text_bygoose(url, re_crawlnumber - 1)
                pass
        return content
        del article
        del url
        del res
        del content
        gc.collect()
