#!/usr/bin/python
# -*- coding: utf-8-*-
# __auther__ : AstroBruce

import requests
import re
import xlrd as xls
import xlwt
from bs4 import BeautifulSoup
import time




#
#
#headers = {'Referer': 'http://www.pss-system.gov.cn/sipopublicsearch/search/searchHomeIndex.shtml',
#		   'User-Agent': 'Mozilla/5.0 (Windows NT 6.3; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Maxthon/4.4.6.2000 Chrome/30.0.1599.101 Safari/537.36',
#
#		   'X-Requested-With': 'XMLHttpRequest'}
#
#postData = {'searchCondition.searchExp': '百度',
#			'searchCondition.dbId': 'VDB',
#			'searchCondition.searchType': 'Sino_foreign',
#			'wee.bizlog.modulelevel': '0200101'}
#data = {}
#session = requests.Session()
#resq = session.post(postUrl, data= postData, headers=headers, timeout=1.0)
#resqSoup = BeautifulSoup(resq.text)
#resqList = resqSoup.findAll(name = 'input', attrs = {'name':re.compile(r'\bvIdHiddenCN\d{12}')})	
#print resqList



def input_xls(xlsName, sheet = 0, col = 0, encoding = 'gbk'):
	xlsData = xls.open_workbook('%s') %xlsName
	xlsTable = xlsData.sheets()[sheet]

	tableRows = xlsTable.nrows

	companyNames = []
	for num in xrange(tableRows):
		companyNames.append(xlsTable.row_values(num)[col].encode(encoding))
	return companyNames

def format_the_post(postUrl, content):
	headers = {'Referer': 'http://www.pss-system.gov.cn/sipopublicsearch/search/searchHomeIndex.shtml',
		   'User-Agent': 'Mozilla/5.0 (Windows NT 6.3; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/46.0.2490.80 Safari/537.36',
		   'X-Requested-With': 'XMLHttpRequest'}

	postData = {'searchCondition.searchExp': '%s' %content,
			'searchCondition.dbId': 'VDB',
			'searchCondition.searchType': 'Sino_foreign',
			'wee.bizlog.modulelevel': '0200101'}
	return postUrl, postData, headers

def session_set(postUrl, postData, postHeaders, timeout = 1.0, times = 3):
	session = requests.Session()
	resq = session.post(postUrl, postData, postHeaders, timeout)
	return resq

def main():	
	postUrl = 'http://www.pss-system.gov.cn/sipopublicsearch/search/smartSearch-executeSmartSearch.shtml'
	headers = {'Referer': 'http://www.pss-system.gov.cn/sipopublicsearch/search/searchHomeIndex.shtml',
		   'User-Agent': 'Mozilla/5.0 (Windows NT 6.3; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/46.0.2490.80 Safari/537.36',
		   'X-Requested-With': 'XMLHttpRequest'}
	xlsData = xls.open_workbook('company.xlsx')
	xlsTable = xlsData.sheets()[0]
	tableRows = xlsTable.nrows
	companyNames = []

	postData = {}
	for num in xrange(tableRows):
		companyNames.append(xlsTable.row_values(num)[0])
		print type(companyNames[num])

	f = xlwt.Workbook(encoding= 'utf-8')
	sheet1 = f.add_sheet(u'sheet1',cell_overwrite_ok=True)
	i = 0 
	for company in companyNames:
		postData = {'searchCondition.searchExp': '%s' %company.encode('utf-8'),
			'searchCondition.dbId': 'VDB',
			'searchCondition.searchType': 'Sino_foreign',
			'wee.bizlog.modulelevel': '0200101'}		
		try:
			session = requests.Session()
			resq = session.post(postUrl, data = postData, headers = headers, timeout = 1.0)
		except Exception, e:
			if hasattr(e, 'code'):
				print 'The server could not fullfill our POST-request, ERROR CODE:', [e.code]
				sheet1.write(i, 0, company.encode('utf-8'))
				sheet1.write(i, 1, 'mistake')
			elif hasattr(e, 'reason'):
				print 'We failed to connect the server,error reason:', e.reason
				sheet1.write(i, 0, company.encode('utf-8'))
				sheet1.write(i, 1, 'mistake')
			else:
				print 'Something wrong that we don\'t know had happened.'
				sheet1.write(i, 0, company.encode('utf-8'))
				sheet1.write(i, 1, 'mistake')
		else:
			resqSoup = BeautifulSoup(resq.text)
			resqList = resqSoup.findAll(name = 'input', attrs = {'name':re.compile(r'\bvIdHiddenCN\d{12}')})	
			if resqList:
				sheet1.write(i, 0, company.encode('utf-8'))
				sheet1.write(i, 1, u'有专利')
				sheet1.write(i, 2, resqList[0]['value'])
			else:
				sheet1.write(i, 0, company.encode('utf-8'))
				sheet1.write(i, 1, u'无专利')	
		i +=1
		print '第%d个公司搜索专利....done' %(i)
		time.sleep(15)
	f.save('output_324.xlsx') 
#-------------------main-----------------
main()
