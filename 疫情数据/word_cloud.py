import openpyxl
from wordcloud import WordCloud
import jieba
import wordcloud

# 读取数据
wb = openpyxl.load_workbook('data.xlsx')
# 获取工作表
ws = wb['国内疫情']
frequency_in = {}
for row in ws.values:
    if row[0] == '省份':
        pass
    else:
        frequency_in[row[0]] = float(row[1])


wordcloud = WordCloud(width=900,  # 词云图片宽度，默认400像素
                        height=383,  # 词云图片高度 默认200像素
                        background_color='black',  # 词云图片的背景颜色，默认为黑色
                        font_path='Microsoft Yahei.ttf',  # 指定字体路径 默认None
                        font_step=1,  # 字号增大的步进间隔 默认1号
                        min_font_size=4,  # 最小字号 默认4号
                        max_font_size=None,  # 最大字号 根据高度自动调节
                        max_words=30,  # 最大词数 默认200
                        scale=15,  # 默认值1。值越大，图像密度越大越清晰
                        prefer_horizontal=0.9,  # 默认值0.90，浮点数类型。表示在水平如果不合适，就旋转为垂直方向，水平放置的词数占0.9%
                        relative_scaling=0,
                        # 默认值0.5，浮点型。设定按词频倒序排列，上一个词相对下一位词的大小倍数。有如下取值：“0”表示大小标准只参考频率排名，“1”如果词频是2倍，大小也是2倍
                        collocations=False,  # 是否包括两个词的搭配
                        mask=None)
# 根据确诊病例的数目生产词云
#wordcloud.generate_from_frequencies(frequency_in)
# 保存词云
#wordcloud.to_file('wordcloud.png')

frequency_out = {}
sheet_name = wb.sheetnames
for each in sheet_name:
    if '洲' in each:
        ws = wb[each]
        for row in ws.values:
            if row[0] == '国家':
                pass
            else:
                frequency_out[row[0]] = float(row[1])

def generate_pic(frequency, name):
    wordcloud = WordCloud(width=900,  # 词云图片宽度，默认400像素
                            height=383,  # 词云图片高度 默认200像素
                            background_color='black',  # 词云图片的背景颜色，默认为黑色
                            font_path='Microsoft Yahei.ttf',  # 指定字体路径 默认None
                            font_step=1,  # 字号增大的步进间隔 默认1号
                            min_font_size=4,  # 最小字号 默认4号
                            max_font_size=None,  # 最大字号 根据高度自动调节
                            max_words=30,  # 最大词数 默认200
                            scale=15,  # 默认值1。值越大，图像密度越大越清晰
                            prefer_horizontal=0.9,  # 默认值0.90，浮点数类型。表示在水平如果不合适，就旋转为垂直方向，水平放置的词数占0.9%
                            relative_scaling=0,
                            # 默认值0.5，浮点型。设定按词频倒序排列，上一个词相对下一位词的大小倍数。有如下取值：“0”表示大小标准只参考频率排名，“1”如果词频是2倍，大小也是2倍
                            collocations=False,  # 是否包括两个词的搭配
                            mask=None)
    # 根据确诊病例的数目生产词云
    wordcloud.generate_from_frequencies(frequency)
    # 保存词云
    wordcloud.to_file('%s.png' %(name))

generate_pic(frequency_in, '国内疫情情况词云图')
generate_pic(frequency_out, '世界疫情情况词云图')
