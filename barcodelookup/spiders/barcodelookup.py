# -*- coding: utf-8 -*-
from scrapy import Spider, Request

import sys
import re, os, requests
from scrapy.utils.response import open_in_browser
from collections import OrderedDict
import time
from xlrd import open_workbook
from shutil import copyfile
import json, re, csv
from scrapy.http import FormRequest
from scrapy.http import TextResponse

def download(url, destfilename):
    if not os.path.exists(destfilename):
        print ("Downloading from {} to {}...".format(url, destfilename))
        try:
            r = requests.get(url, stream=True)
            with open(destfilename, 'wb') as f:
                for chunk in r.iter_content(chunk_size=1024):
                    if chunk:
                        f.write(chunk)
                        f.flush()
        except:
            print ("Error downloading file.")

def readExcel(path):
    wb = open_workbook(path)
    result = []
    for sheet in wb.sheets():
        number_of_rows = sheet.nrows
        number_of_columns = sheet.ncols
        herders = []
        for row in range(0, number_of_rows):
            values = OrderedDict()
            for col in range(number_of_columns):
                value = (sheet.cell(row,col).value)
                if row == 0:
                    herders.append(value)
                else:

                    values[herders[col]] = value
            if len(values.values()) > 0:
                result.append(values)
        break

    return result


class AngelSpider(Spider):
    name = "barcodelookup"
    start_urls = 'https://www.barcodelookup.com/'
    count = 0
    use_selenium = False
    site_name = "TropicMarket"
    ean_codes = readExcel("sample_EAN.xlsx")
    models = []
    headers = ['EAN', 'image1', 'image2', 'image3', 'image4', 'image5','image6']
    for code in ean_codes:
        item = OrderedDict()
        item['EAN'] = str(int(code['EAN']))
        models.append(item)

    def start_requests(self):
        for i, val in enumerate(self.models):
            ern_code = val['EAN']
            yield Request(self.start_urls + ern_code, callback=self.parse1, meta={'ean':ern_code})

    def parse1(self, respond):
        item = OrderedDict()
        for i, key in enumerate(self.headers):
            item[key] = ''

        item['EAN'] = respond.meta['ean']
        picture_urls = respond.xpath('//div[@id="thumbs"]/div/img/@src').extract()
        if picture_urls:
            pdf_path = "images/{}/".format(item['EAN'])
            if not os.path.exists(pdf_path):
                os.makedirs(pdf_path)
        for i, picture_url in enumerate(picture_urls):
            filename = item['EAN'] + "_" + str(i+1) + "_" + self.site_name + ".jpg"
            item['image'+str(i+1)] = filename
            download(picture_url, pdf_path+filename)



