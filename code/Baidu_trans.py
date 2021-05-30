import http.client
import hashlib
import urllib
import random
import json
import time


class Baidu_trans():
    def __init__(self):
        self.keyword = ""

    def translate_baidu(self, keyword):
        appid = '20210527000844652'  # 填写你的appid
        secretKey = 'd6hfYWZodxqDLm5g0pC3'  # 填写你的密钥

        httpClient = None
        myurl = '/api/trans/vip/translate'

        fromLang = 'auto'  # 原文语种
        toLang = 'en'  # 译文语种
        salt = random.randint(32768, 65536)
        q = keyword
        sign = appid + q + str(salt) + secretKey
        sign = hashlib.md5(sign.encode()).hexdigest()
        myurl = myurl + '?appid=' + appid + '&q=' + urllib.parse.quote(
            q) + '&from=' + fromLang + '&to=' + toLang + '&salt=' + str(
            salt) + '&sign=' + sign

        try:
            time.sleep(1)       # 睡眠1秒
            httpClient = http.client.HTTPConnection('api.fanyi.baidu.com')
            httpClient.request('GET', myurl)

            # response是HTTPResponse对象
            response = httpClient.getresponse()
            result_all = response.read().decode("utf-8")
            result = json.loads(result_all)

            return result["trans_result"][0]["dst"]

        except Exception as e:
            print(e)
        finally:
            if httpClient:
                httpClient.close()


# br = Baidu_trans()
# results = br.translate("枣庄市峄城区阴平镇发展绿色生态循环农业，助推乡村产业振兴")
# print(results)
