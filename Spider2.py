"""
    淘宝评论爬虫(beautifulsoup)
"""
# -*- coding: utf-8 -*-
import re
import time
import random
import pymongo
import numpy as np
from datetime import datetime
from bs4 import BeautifulSoup
from selenium import webdriver
from wordcloud import WordCloud
from pyecharts import options as opts
from pyecharts.globals import ThemeType
from pyecharts.charts import Bar, Pie, Page
from selenium.webdriver.support.wait import WebDriverWait


search_word = '手机'  # 搜索词
page_num = 5  # 抓取页数
collection_name = 'testphone'  # 集合名
global img_num  # 有图评价数量
global vip_num  # 超级会员数
global comment_num  # 总评价数
global append_num  # 追加评价数
global content
img_num = vip_num = comment_num = append_num = 0

def elementlist():# 构造元素列表
    i, j = 1, 6
    elementlist = []
    while i <= page_num:
        if j < 11:
            elementlist.append(j)
            i += 1
            j += 1
        else:
            elementlist.append(j)
            i += 1
    return elementlist

def set_ip_proxy():
    pass

def request_pagesource():    # 请求网页
    global content
    # url = 'https://www.taobao.com'
    login_url = 'https://login.taobao.com'
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_experimental_option('excludeSwitches', ['enable-automation'])# 开发者模式
    browser = webdriver.Chrome(options=chrome_options)  # 浏览器对象初始化并赋给browser对象
    WebDriverWait(browser, 10)
    browser.get(login_url)  # 打开淘宝网页
    browser.maximize_window()  # 窗口最大化
    time.sleep(20)  # 等待20秒，用手机淘宝扫码实现登录
    try:
        try:
            browser.find_element_by_css_selector('#q').send_keys(search_word)  # 使用css选择器找到淘宝页面的输入框并输入手机
        except Exception:
            time.sleep(5)
    except Exception:
        print('登录超时，请重新登录')
    time.sleep(random.randint(3, 9))
    browser.find_element_by_class_name('btn-search').click()  # 点击搜索按钮
    time.sleep(random.randint(3, 9))
    browser.find_element_by_link_text('销量').click()
    time.sleep(random.randint(3, 9))
    pagesource = browser.page_source    # 获取pagesource
    soup = BeautifulSoup(pagesource, 'lxml')
    item_id = soup.find('div', class_="item J_MouserOnverReq").find('img')['id']   # 找到商品页的第一个商品
    browser.find_element_by_css_selector('#'+item_id).click()
    current_window = browser.current_window_handle   # 获取到当前所有的句柄
    all_handles = browser.window_handles   # 所有的句柄存放在列表当中
    # print(all_handles)  #打印句柄
    '''获取非最初打开页面的句柄'''
    for handles in all_handles:
        if current_window != handles:
            browser.switch_to.window(handles)
    # currenturl = browser.current_url
    # print(currenturl)
    browser.find_element_by_id('J_ItemRates').click()
    time.sleep(random.randint(3, 9))
    for i in elementlist():
        time.sleep(random.randint(3, 9))
        html = browser.page_source
        result = re.findall('<div id="J_Reviews".*<div class="tm-trypage">', html, re.S)
        # print(result)#result是一个列表
        result = "".join(result)#将result转为str
        data_processing(result)
        time.sleep(random.randint(3, 9))
        nextpage = browser.find_element_by_css_selector('#J_Reviews > div > div.rate-page > div > a:nth-child(%i)' % i)
        # nextpages = browser.find_element_by_link_text('下一页>>')
        browser.execute_script('arguments[0].click();', nextpage)
    print('Finish')
    content = '总评论数:' + str(comment_num) + '\t追加评论数:' + str(append_num) + '\n超级会员数:' + str(vip_num) + '\t有图评论数:' + str(img_num)
    print(content)
    browser.quit()

def data_processing(result):   # 数据处理
    global img_num
    global vip_num
    global comment_num
    global append_num
    soup = BeautifulSoup(result, 'lxml')# lxml HTML 解析器
    tr = soup.find_all('tr')# 找到所有tr标签
    datalist = []
    for i in tr:
        try:
            dic = {}
            dic['是否为追加评论'] = 1
            dic['评论日期'] = i.find('div', class_="tm-rate-date").text
            dic['评论'] = i.find('div', class_="tm-rate-premiere").find('div', class_='tm-rate-fulltxt').text
            dic['追加评论时间'] = i.find('span', class_='tm-rate-daydiff').text.rstrip('：')
            dic['追加评论'] = i.find('div', class_="tm-rate-append").find('div', class_='tm-rate-fulltxt').text
            dic['商品详情'] = i.find('div', class_="rate-sku").text
            dic['买家姓名'] = i.find('div', class_="rate-user-info").text
            if (i.find('p', class_="gold-user") is None):
                dic['是否为超级会员'] = 0
            else:
                dic['是否为超级会员'] = 1
                vip_num += 1
            if (i.find('li') is None):
                dic['是否有图'] = 0
            else:
                dic['是否有图'] = 1
                img_num += 1
                img_list = []
                li_list = (i.find_all('li'))
                for img in li_list:
                    img_list.append(img['data-src'])  # 找到li标签的data-src属性
                dic['图片'] = img_list
            append_num += 1
            datalist.append(dic)
        except AttributeError:
            dic = {}
            dic['是否为追加评论'] = 0
            dic['评论日期'] = i.find('div', class_="tm-rate-date").text
            dic['评论'] = i.find('div', class_="tm-rate-content").find('div', class_="tm-rate-fulltxt").text
            dic['追加评论时间'] = ''
            dic['追加评论'] = ''
            dic['商品详情'] = i.find('div', class_="rate-sku").text
            dic['买家姓名'] = i.find('div', class_="rate-user-info").text
            if (i.find('p', class_="gold-user") is None):
                dic['是否为超级会员'] = 0
            else:
                dic['是否为超级会员'] = 1
                vip_num += 1
            if (i.find('li') is None):
                dic['是否有图'] = 0
            else:
                dic['是否有图'] = 1
                img_num += 1
                img_list = []
                li_list = (i.find_all('li'))
                for img in li_list:
                    img_list.append(img['data-src'])
                dic['图片'] = img_list
            datalist.append(dic)
        finally:
            comment_num += 1
    database(datalist)

def make_wordcloud():  # 做词云图
    client = pymongo.MongoClient("mongodb://localhost:27017/")  # 连接MongoDB数据库
    db = client['testcomment']  # 建立testcomment数据库
    col = db[collection_name]  # 建立 collection
    comment_list = []
    for data in col.find({}, {"_id": 0, "评论": 1, "追加评论": 1}):  # 从数据库里查找所有评论数据
        comment_list.append(data['评论'])
        comment_list.append(data['追加评论'])
    text = str(comment_list)
    stopword = {'拍照效果': 0, '显示效果': 0, '电池续航': 0, '其他特色': 0, '通信音质': 0, '运行速度': 0, '此用户没有填写评论': 0}  # 停用词，词云图中会排除这些词
    wordcloud = WordCloud(font_path="simhei.ttf", width=1000, height=800, stopwords=stopword).generate(text)
    image = wordcloud.to_image()
    image.show()
    print('词云图已生成')

def database(datalist):
    client = pymongo.MongoClient("mongodb://localhost:27017/")  # 连接MongoDB数据库
    db = client['testcomment']
    col = db[collection_name]
    try:
        col.insert_many(datalist)
    except TypeError as error:  # 如果没爬到数据datalist为空，会抛出TypeError
        print('Error:', error, '请重新操作')

def analyse_data():
    client = pymongo.MongoClient("mongodb://localhost:27017/")
    db = client['testcomment']
    col = db[collection_name]
    result = col.find({}, {'_id': 0, '评论日期': 1, "评论": 1, "追加评论": 1})
    comment_list = []
    date_list = []
    date_dic = {}
    for date in result:  # 把日期和评论字数放入列表中
        date_list.append(date['评论日期'])
        comment_list.append(len(date['评论']))
        if len(date['追加评论']) != 0:
            comment_list.append(len(date['追加评论']))
    for date in date_list:  # 统计日销量
        if date in date_dic:
            date_dic[date] += 1
        else:
            date_dic[date] = 1
    # date_list = list(date_dic)
    data_dic = {}
    arr = np.array(list(date_dic))  # 给日期排序
    date_list = arr[np.argsort([datetime.strptime(i, '%m.%d') for i in list(date_dic)])].tolist()
    for i in date_list:
        data_dic[i] = date_dic[i]
    # s = pd.Series(data_dic)  # series表示一维数组
    # comment_list.sort()
    word_num_dic = {'0-10': 0, '11-20': 0, '21-30': 0, '31-40': 0, '41-50': 0, '51-60': 0,
                    '61-70': 0, '71-80': 0, '91-100': 0, '101-110': 0, '111-120': 0, '121-130': 0,
                    '131-140': 0, '141-150': 0, '151-160': 0, '161-170': 0, '171-180': 0, '181-190': 0,
                    '191-200': 0, '201-300': 0}
    for i in word_num_dic:  # 统计字数
        for j in comment_list:
            if j >= int(i.split('-')[0]) and j <= int(i.split('-')[-1]):
                word_num_dic[i] += 1
    # print(ndic)
    bar1 = (  # 生成柱状图
        Bar()
        .add_xaxis(list(data_dic.keys()))
        .add_yaxis('商品', list(data_dic.values()))
        .set_global_opts(  # 设置全局参数
            title_opts=opts.TitleOpts(title='商品日销量'),
            yaxis_opts=opts.AxisOpts(splitline_opts=opts.SplitLineOpts(is_show=True)),  # 添加纵坐标标线
            datazoom_opts=[opts.DataZoomOpts(range_start=10, range_end=80, is_zoom_lock=False)]  # 添加滑动条
        )
    )
    bar1.chart_id = '6f180e81787a48539de004c1eb847c1e'
    bar2 = (  # 生成柱状图
        Bar({'theme': ThemeType.MACARONS})
        .add_xaxis(list(word_num_dic.keys()))
        .add_yaxis('商品', list(word_num_dic.values()))
        .set_global_opts(  # 设置全局参数
            title_opts=opts.TitleOpts(title='字数统计'),
            yaxis_opts=opts.AxisOpts(splitline_opts=opts.SplitLineOpts(is_show=True)),
            datazoom_opts=[opts.DataZoomOpts(range_start=10, range_end=80, is_zoom_lock=False)]
        )
    )
    bar2.chart_id = '9a412341322b4d92bba49e3fd932a7f0'
    pie1 = (  # 生成柱状图
        Pie()
        .add(series_name='用户身份', data_pair=[('超级会员', vip_num), ('普通用户', comment_num-vip_num)])
        .set_global_opts(  # 设置全局参数
            title_opts=opts.TitleOpts(title='超级会员比例'),
        )
    )
    pie1.chart_id = '7ed5d413532646909f127fd36f584081'
    pie2 = (  # 生成柱状图
        Pie(init_opts=opts.InitOpts(theme=ThemeType.WESTEROS))
        .add(series_name='有无图片', data_pair=[('有图评价', img_num), ('无图评价', comment_num-img_num)])
        .set_global_opts(  # 设置全局参数
            title_opts=opts.TitleOpts(title='有图评价比例'),
        )
    )
    pie2.chart_id = 'aa3f038854e144dbaa6e9892d4ca31fc'
    pie3 = (  # 生成柱状图
        Pie(init_opts=opts.InitOpts(theme=ThemeType.INFOGRAPHIC))
        .add(series_name='是否追加', data_pair=[('追加评价', append_num), ('普通评价', comment_num-append_num)])
        .set_global_opts(  # 设置全局参数
            title_opts=opts.TitleOpts(title='追加评价比例'),
        )
    )
    pie3.chart_id = '235863de6fe74a2aa9273e30ed4a7f66'
    page = Page(page_title='评论分析报告')
    page.add(bar1, bar2, pie1, pie2, pie3)
    page.render()
    Page.save_resize_html(cfg_file=r'chart_config.json')
    with open('resize_render.html', 'r+', encoding='utf-8') as f:
        f.seek(573)
        html = f.read()
        f.seek(573)
        f.write('   <h1>评价分析结果：</h1>\n    <p>'+content+'</p>\n ')
        f.write(html)
    print(r'已完成分析   请打开E:\Python\learnPython\testspider\FinalProject\resize_render.html查看分析结果')

def main():
    request_pagesource()
    make_wordcloud()
    analyse_data()

if __name__ == '__main__':
    main()
