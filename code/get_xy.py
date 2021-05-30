#coding=gbk
import sys
import io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='gb18030')
# ��δ�������þ��ǰѱ�׼�����Ĭ�ϱ����޸�Ϊgb18030��Ҳ������cmd��ʾ����GBK��ͬ��

from fake_useragent import UserAgent
import openpyxl
import json
import requests

ua = UserAgent()

def get_xy(place):
    url = 'http://api.map.baidu.com/geocoding/v3/?address=' +place+ '&output=json&ak={#��Ҫ���Լ���ak}'
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