# -*- coding: utf-8 -*-
#writtenby zcc 2021/1/11
import urllib
import urllib.request
import os
import requests
import mechanize
import datetime
from selenium import webdriver
from bs4 import BeautifulSoup #解析库
import re
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
import time
#通过程序调用浏览器
#这边我使用的是chrome 需要根据你自己的浏览器的版本到 https://npm.taobao.org/mirrors/chromedriver/下载chromedriver.exe
#并放到和python.exe 同一目录下 https://www.cnblogs.com/lfri/p/10542797.html 注意自己浏览器的版本
#如果是其他浏览器不用chrome的话可以百度下流程
def get_browser():
    # chrome_options = Options()
    # chrome_options.add_argument("--headless")
    # br = webdriver.Chrome(options=chrome_options)
    br = webdriver.Chrome()
    return br

def check_path():
    if not os.path.exists(save_path):
        os.mkdir(save_path)
    for i in range(len(company)):
        temp_path = os.path.join(save_path, company[i])
        path.append(temp_path)
        if not os.path.exists(temp_path):
            os.mkdir(temp_path)
#根据指定的文件夹路径和获取的当前页面中的url的地址下载
#文件到对应的文件夹(这边是按照公司的股票号码和公司名称)
def download_pdf_file(save_path,url_pdf):
    filename = url_pdf.split('/')[-1]
    filepath = os.path.join(save_path,filename)
    #如果文件不存在就下载
    if not os.path.exists(filepath):
        try:
            urllib.request.urlretrieve(url_pdf, filename=filepath)
        except Exception as e:
            print("Error occurred when downloading file, error message:")
            print(e)

def set_time(browser):
    #先设置时间 不同公司的都是相同的可以统一设置时间
    start_time='2000-01-01'
    end_time='2020-12-31'
    start_date = datetime.date(*map(int, start_time.split('-'))).strftime("%Y/%m/%d")
    end_date = datetime.date(*map(int, end_time.split('-'))).strftime("%Y/%m/%d")
    js = "document.getElementById('searchDate-From').value='{}';".format(start_date)
    je = "document.getElementById('searchDate-To').value='{}';".format(end_date)
    browser.execute_script(js)
    browser.find_element(By.ID, "searchDate-From").click()
    browser.execute_script(je)
    browser.find_element(By.ID, "searchDate-To").click()
    #在网页200，300的位置点击下确定选择的时间,相当于人工的交互
    time.sleep(1)
# def set_annual_report(browser):

def pdf_down(soup,save_path,company_name):
    # print(soup)
    link_tags = []
    tmps = soup.find_all('a')
    for u in tmps:
        try:
            link=u['href']
            if link[-4:]==".pdf" and  link.find('/listedco/listconews/')==0:
                dw_link='http://www.hkexnews.hk'+link
                link_tags.append(dw_link)
        except:
            continue
    for i in range(len(link_tags)):
        print(link_tags[i])
        save_pdf_path=os.path.join(save_path,company_name)
        download_pdf_file(save_pdf_path,link_tags[i])

#设置年报的选项的按钮的选择
def set_annual_report(browser):
    #找到所有的标签 点击 这个查找是查看网页的html结构获得的
    browser.find_element_by_xpath('//li[@class="filter__container-input searchDocType"]/div/div[@class="combobox-group-container clearfix"]/div/div/div/div/a').click()  # 先点击下拉按钮
    browser.find_element_by_link_text("標題類別").click()
    browser.find_element_by_xpath('//li[@class="filter__container-input searchDocType"]/div/div[@class="combobox-group-container clearfix"]/div[@class="tier1-wrap searchType-Categroy filter__dropdown-js"]/div[@id="rbAfter2006"]/div/div/div/a').click()
    browser.find_element_by_link_text("財務報表/環境、社會及管治資料").click()
    browser.find_element_by_link_text("年報").click()

def choose_company(company_number,browser):
    browser.find_element_by_id('searchStockCode').send_keys(company_number)
    time.sleep(1) #留时间给浏览器进行反应
    #默认选择搜索到的第一个。
    # browser.find_element_by_xpath('//li[@class="filter__container-input searchStockCodeName"]/div/div[@class="autocomplete-group"]/div[@id="autocomplete-list-0"]/div[@class="slimScrollDiv"]/div/table/tbody/tr').click()
    table = browser.find_element_by_xpath('//li[@class="filter__container-input searchStockCodeName"]/div/div[@class="autocomplete-group"]/div[@class="autocomplete-suggestions"]/div/div/table/tbody/tr')
    time.sleep(1)
    table.click()
    time.sleep(1)

def down_file(save_path,url,company):
    '''
    :param save_path: 文件下载本地地址
    :param url:表单查询地址
    :param company: 需要查询的公司列表
    :param url_table: 需要查询的公司列表在此链接中进行选择
    :return:
    '''
    '''
    查询条件设置 現有上市證券 股份代號/股份名稱 標題類別 年報 開始日期2000/01/01  完結日期2020/12/31 搜尋
    '''
    for i in range(len(company)):
        browser = get_browser()
        browser.get(url)
        #选择上市公司代码
        choose_company(company[i], browser)
        #设置选择年报
        set_annual_report(browser)
        #设置查找时间
        set_time(browser)
        # 当设置完成后进行搜索的动作
        browser.find_element_by_link_text("搜尋").click()
        # 等待三秒 看网速情况 是为了加载完成后需要等待页面加载好再进行后后续处理
        time.sleep(1)
        page = browser.page_source
        soup = BeautifulSoup(page, 'html.parser')
        pdf_down(soup,save_path,company[i])
        print(company[i],"年报下载完成")
        browser.close()
if __name__ == '__main__':
    #下载设置
    url ='https://www1.hkexnews.hk/search/titlesearch.xhtml?lang=zh'
    # url = 'https://www.hkexnews.hk/index_c.htm'
    # 可以把需要添加公司的代码加上
    company = ['00001', '00002']
    # 用于保存对应公司的名称
    path = []
    # 创建有个专门用于保存文件路径的路径
    save_path = './PDFFile'
    #检查文件夹没有的就进行创建
    check_path()
    #设置查询条件并进行下载
    down_file(url=url,company=company,save_path=save_path)




