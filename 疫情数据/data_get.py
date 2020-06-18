import requests
import re
import json
from lxml import etree


class Get_data():
    # 获取数据
    def get_data(self):
        response = requests.get('https://voice.baidu.com/act/newpneumonia/newpneumonia')
        with open('html.txt', 'w') as file:
            file.write(response.text)

    # 提取更新时间
    def get_time(self):
        with open('html.txt', 'r') as file:
            text = file.read()
        time = re.findall('"mapLastUpdatedTime":"(.*?)"', text)[0]
        return time

    # 解析数据
    def parse_data(self):
        with open('html.txt', 'r') as file:
            text = file.read()
        html = etree.HTML(text)
        result = html.xpath('//script[@type="application/json"]/text()')
        result = result[0] # json.loads()方法可以将字符串转化为python数据类型
        result = json.loads(result)
        result = result['component'][0]['caseList']
        result = json.dumps(result)
        with open('data.json', 'w') as file:
            file.write(result)
            print('数据已写入json文件……')

