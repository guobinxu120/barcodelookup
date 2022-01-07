# -*- coding: utf-8 -*-

# Define your item pipelines here
#
# Don't forget to add your pipeline to the ITEM_PIPELINES setting
# See: https://doc.scrapy.org/en/latest/topics/item-pipeline.html
from scrapy import signals
import xlsxwriter
import os

class BarcodelookupPipeline(object):
    def __init__(self):
            #Instantiate API Connection
            self.files = {}


    @classmethod
    def from_crawler(cls, crawler):
        pipeline = cls()
        crawler.signals.connect(pipeline.spider_opened, signals.spider_opened)
        crawler.signals.connect(pipeline.spider_closed, signals.spider_closed)
        return pipeline


    def spider_opened(self, spider):
        pass

    def spider_closed(self, spider):
        pass
        filepath = 'output.xlsx'
        if os.path.isfile(filepath):
            os.remove(filepath)
        workbook = xlsxwriter.Workbook(filepath)
        sheet = workbook.add_worksheet('output.xlsx')
        data = spider.models
        heades = spider.headers
        flag =True
        headers = []
        for index, value in enumerate(data):
            if flag:
                for col, val in enumerate(headers):
                    headers.append(val)
                    sheet.write(index, col, val)
                flag = False
            for col, key in enumerate(headers):
                sheet.write(index+1, col, value[key])
        workbook.close()
