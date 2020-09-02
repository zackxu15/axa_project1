import scrapy
import glob
from scrapy import Selector
import codecs
import pandas as pd
import os 
import time



# html = 'G:\Job\AXA-Equitable\Test1_source\ZoomInfo 2.0 (2020-07-07 4_58_50 PM).html'
# f = codecs.open(html, 'r', 'utf-8').read()
# sel = Selector(text = f)

# print(sel.xpath(
#     '//zi-row-text/span[contains(@class,"text")]//text()').extract_first())
# print(sel.xpath(
#     '//zi-dotten-text[contains(@class,"company-name-link")]/div/div//text()').extract_first())
# print(sel.xpath(
#     '//zi-dotten-text[@class="person-title"]/div/div//text()').extract_first())
# print(sel.xpath('//a[@class="email-link"]/@title').extract_first())
# print(
#     sel.xpath('//zi-text/a[contains(@class,"record-data")]//text()').getall())

def info_create(path, num):
    html = codecs.open(path, 'r', 'utf-8').read()
    sel = Selector(text=html)
    name = sel.xpath(
        '//zi-row-text/span[contains(@class,"text")]//text()').extract_first()
    company = sel.xpath(
        '//zi-dotten-text[contains(@class,"company-name-link")]/div/div//text()').extract_first()
    postion = sel.xpath(
        '//zi-dotten-text[@class="person-title"]/div/div//text()').extract_first()
    email = sel.xpath('//a[@class="email-link"]/@title').extract_first()
    phones = sel.xpath(
        '//zi-text/a[contains(@class,"record-data")]//text()').getall()

    tele = {}
    i=1
    for phone in phones:
        tele["phone{0}".format(i)] = phone
        i += 1

    info = {
        'name':name,
        'company': company,
        'postion': postion,
        'email': email,
    }
    info.update(tele)
    info = {num : info}
    return info
    
def data_create(files):
    num = 0
    dic = {}
    for file in files:
        num += 1
        info = info_create(file, num)
        if dic == None:
            dic = info
        else:
            dic.update(info)

    data = pd.DataFrame.from_dict(dic, orient='index').drop_duplicates().dropna(subset=['name'])
    data.reset_index(drop=True, inplace=True)
    return data

file_path = os.getcwd()
files = glob.glob(file_path+r"\*.html")
#files = glob.glob("G:\Job\AXA-Equitable\Test1_source\*.html")

data = data_create(files)

timestr = time.strftime("%Y%m%d")

data.to_excel('data_'+timestr+'.xlsx', index=False)

