# -*- coding: utf-8 -*-
import scrapy
import json
import os
from scrapy.shell import inspect_response
from zipfile import ZipFile

import xlrd
import openpyxl
import re
from openpyxl import Workbook
from openpyxl import load_workbook


class BiduiSpider(scrapy.Spider):
    name = 'biduipingtai_spider'
    allowed_domains = ["http://132.121.80.158:8090/"]

    def start_requests(self):
        url = 'http://132.121.80.158:8090/plversion/security_check'
        headers = {
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
            'Host': '132.121.80.158:8090',
            'Origin': 'http://132.121.80.158:8090',
            'Referer': 'http://132.121.80.158:8090/plversion/com.gxlu.security.view.login.d',
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/55.0.2883.87 Safari/537.36',
            'Upgrade-Insecure-Requests': '1'
        }
        login_info = {
            'PL': '1',
            'password': '654321',
            'username': 'tool_admin'
        }
        yield scrapy.FormRequest(
            url=url,
            method='POST',
            headers=headers,
            formdata=login_info,
            callback=self.parse_device,
            dont_filter=True)

    # 获取下载设备比对结果的路径
    def parse_device(self, response):
        url = 'http://132.121.80.158:8090/plversion/dorado/view-service'
        body = '''{"action":"remote-service","service":"outputAllDataService#findAllDataOutPutExcel","parameter":{"pageflag":"TMP_IPRANNETL2WGDIFFERENT","pageNo":1,"pageSize":60000},"context":{"orgId":null,"DataORG_ID":null,"pageflag":"TMP_IPRANNETL2WGDIFFERENT","ListStatisSQL":null,"Createdate":null,"CompareStatus":null,"viewId":"com.gxlu.statisticommon.view.charttest"},"loadedDataTypes":["dataTypeDatatable","dataTypeGisProject","dataTypeTime","dataTypeTelant","dataTypeOrgStruct"]}'''
        headers = {
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
            'Host': '132.121.80.158:8090',
            'Origin': 'http://132.121.80.158:8090',
            'Referer': 'http://132.121.80.158:8090/plversion/com.gxlu.statisticommon.view.ListConfig.d?viewId=com.gxlu.statisticommon.view.charttest&pageflag=TMP_IPRANNETL2WGDIFFERENT&wordUri=4GdataCompare',
            'Content-Type': 'text/javascript',
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/55.0.2883.87 Safari/537.36'
        }
        yield scrapy.Request(
            url=url,
            method='POST',
            body=body,
            headers=headers,
            callback=self.parse_device_download,
            dont_filter=True
        )

    # 获取下载文件名
    def parse_device_download(self, response):
        text = response.text
        text = json.loads(text)
        filename = text['data']
        url = 'http://132.121.80.158:8090/plversion/excel/' + filename
        yield scrapy.Request(
            url=url,
            method='GET',
            callback=self.parse_card,
            dont_filter=True
        )

    # 下载设备比对结果，获取下载板卡比对结果的路径
    def parse_card(self, response):
        # inspect_response(response, self)
        filename_device = os.path.join(os.path.abspath('.'), '设备.zip')
        with open(filename_device, 'wb') as f:
            f.write(response.body)
        url = 'http://132.121.80.158:8090/plversion/dorado/view-service'
        body = '''{"action":"remote-service","service":"outputAllDataService#findAllDataOutPutExcel","parameter":{"pageflag":"TMP_IPRANCARDTL2WGDIFFERENT","pageNo":3,"pageSize":60000},"context":{"orgId":null,"DataORG_ID":null,"pageflag":"TMP_IPRANCARDTL2WGDIFFERENT","ListStatisSQL":null,"Createdate":null,"CompareStatus":null,"viewId":"com.gxlu.statisticommon.view.charttest"},"loadedDataTypes":["dataTypeDatatable","dataTypeOrgStruct","dataTypeTelant","dataTypeGisProject","dataTypeTime"]}'''
        headers = {
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
            'Host': '132.121.80.158:8090',
            'Origin': 'http://132.121.80.158:8090',
            'Referer': 'http://132.121.80.158:8090/plversion/com.gxlu.statisticommon.view.ListConfig.d?viewId=com.gxlu.statisticommon.view.charttest&pageflag=TMP_IPRANCARDTL2WGDIFFERENT&wordUri=4GdataCompare',
            'Content-Type': 'text/javascript',
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/55.0.2883.87 Safari/537.36'
        }
        yield scrapy.Request(
            url=url,
            method='POST',
            body=body,
            headers=headers,
            callback=self.parse_card_download,
            dont_filter=True
        )
        # inspect_response(response, self)

    def parse_card_download(self, response):
        text = response.text
        text = json.loads(text)
        filename = text['data']
        url = 'http://132.121.80.158:8090/plversion/excel/' + filename
        yield scrapy.Request(
            url=url,
            method='GET',
            callback=self.parse_circuit,
            dont_filter=True
        )

    def parse_circuit(self, response):
        filename_card = os.path.join(os.path.abspath('.'), '板卡.zip')
        with open(filename_card, 'wb') as f:
            f.write(response.body)

        url = 'http://132.121.80.158:8090/plversion/dorado/view-service'
        body = '''{"action":"remote-service","service":"outputAllDataService#findAllDataOutPutExcel","parameter":{"pageflag":"TMP_IPRANLINKTL2WGDIFFERENT","pageNo":2,"pageSize":60000},"context":{"orgId":null,"DataORG_ID":null,"pageflag":"TMP_IPRANLINKTL2WGDIFFERENT","ListStatisSQL":null,"Createdate":null,"CompareStatus":null,"viewId":"com.gxlu.statisticommon.view.charttest"},"loadedDataTypes":["dataTypeGisProject","dataTypeOrgStruct","dataTypeDatatable","dataTypeTime","dataTypeTelant"]}'''
        headers = {
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
            'Host': '132.121.80.158:8090',
            'Origin': 'http://132.121.80.158:8090',
            'Referer': 'http://132.121.80.158:8090/plversion/com.gxlu.statisticommon.view.ListConfig.d?viewId=com.gxlu.statisticommon.view.charttest&pageflag=TMP_IPRANLINKTL2WGDIFFERENT&wordUri=4GdataCompare',
            'Content-Type': 'text/javascript',
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/55.0.2883.87 Safari/537.36'
        }
        yield scrapy.Request(
            url=url,
            method='POST',
            body=body,
            headers=headers,
            callback=self.parse_circuit_download,
            dont_filter=True
        )

    def parse_circuit_download(self, response):
        text = response.text
        text = json.loads(text)
        filename = text['data']
        url = 'http://132.121.80.158:8090/plversion/excel/' + filename
        yield scrapy.Request(
            url=url,
            method='GET',
            callback=self.parse_success,
            dont_filter=True
        )

    def parse_success(self, response):
        filename_circuit = os.path.join(os.path.abspath('.'), '电路.zip')
        with open(filename_circuit, 'wb') as f:
            f.write(response.body)
        size = os.path.getsize(filename_circuit)
        print(size)

        # 对下载的压缩包进行解压，合并Excel表
        choice = input('是否需要解压并整合文件？1：是 2：否  请选择：')
        if choice == '1':
            device_Excel_list = self.unzip('设备.zip')
            card_Excel_list = self.unzip('板卡.zip')
            circuit_Excel_list = self.unzip('电路.zip')

            self.combination(device_Excel_list, '设备')
            self.combination(card_Excel_list, '板卡')
            self.combination(circuit_Excel_list, '电路')
        else:
            pass

    def unzip(self, filename):
        f = ZipFile(filename)
        name_list = f.namelist()
        f.extractall()
        return name_list

    def combination(self, filelist, filename_new):
        ILLEGAL_CHARACTERS_RE = re.compile(
            r'[\000-\010]|[\013-\014]|[\016-\037]')
        num_of_file = len(filelist)
        if num_of_file == 1:
            os.rename(filelist[0], filename_new + '.xls')
        else:
            wb = Workbook()
            ws = wb.active
            count = 0
            for file in filelist:
                print('正在读取' + str(file) + '...')
                data = xlrd.open_workbook(file)
                print('读取完毕！')
                table = data.sheet_by_index(0)
                nrows = table.nrows

                ws.append(table.row_values(0))
                for row in range(1, nrows):
                    try:
                        ws.append(table.row_values(row))
                        count += 1
                        print('已经复制' + str(count) + '条数据！')
                    except openpyxl.utils.exceptions.IllegalCharacterError:
                        tmp = table.row_values(row)
                        for i in range(len(tmp)):
                            tmp[i] = ILLEGAL_CHARACTERS_RE.sub(r'', tmp[i])
                        ws.append(tmp)
                        count += 1
                        print('已经复制' + str(count) + '条数据！')
            wb.save(filename_new + '.xlsx')
