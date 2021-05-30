#coding=gbk
import sys
import io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='gb18030')
# 这段代码的作用就是把标准输出的默认编码修改为gb18030，也就是与cmd显示编码GBK相同。

from fake_useragent import UserAgent
import openpyxl
import json
import requests

ua = UserAgent()

def get_xy(place):
    url = 'http://api.map.baidu.com/geocoding/v3/?address=' +place+ '&output=json&ak={#需要用自己的ak}'
    headers = {
        'User-Agent': ua.random
    }

    try:
        res = requests.get(url, headers=headers)
        lng = res.json()['result']['location']['lng']
        lat = res.json()['result']['location']['lat']

        return lng, lat
    except:
        return None
        pass