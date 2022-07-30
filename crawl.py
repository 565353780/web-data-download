from lxml import etree
import requests
import csv
import time
import random

class Company:
	def __init__(self):
		self.name = ""
		self.tel = ""
		self.email = ""
		self.website = ""
		self.adress = ""
		self.simpleinfo = ""
		self.orgname = ""
	def __str__(self):
		return "{},{},{},{},{},{},{}".format(self.name, self.tel, self.email, self.website, self.adress, self.simpleinfo, self.orgname)

	def toArray(self):
		return [self.name, self.tel, self.email, self.website, self.adress, self.simpleinfo, self.orgname]

header = {
	'Host': 'www.tianyancha.com',
	'Connection': 'keep-alive',
	'Cache-Control': 'max-age=0',
	'Upgrade-Insecure-Requests': '1',
	'User-Agent': r'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/70.0.3538.102 Safari/537.36 Edge/18.18362',
	'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8',
	'Referer':  'https://www.tianyancha.com/company/',
	'Accept-Encoding': 'gzip, deflate, br',
	'Accept-Language': 'zh-CN,zh;q=0.9',
	'Cookie': r'Cookie: bannerFlag=undefined; cloud_token=1147bdc4820d4ec78f5a20b0e965936e; Hm_lpvt_e92c8d65d92d534b0fc290df538b4758=1575708094; __insp_wid=677961980; auth_token=eyJhbGciOiJIUzUxMiJ9.eyJzdWIiOiIxMzcyMDEwODExNyIsImlhdCI6MTU3NTcwMzU2MywiZXhwIjoxNjA3MjM5NTYzfQ.1eSbT0GCR594Dt4CS1EkDgOK7wVLuBFWQbvemuxxAIlo5CyFSXj1bxY_Vvvnbn2Kawv-H4j_sS9nZ7tL-TO65w; __insp_targlpu=aHR0cHM6Ly93d3cudGlhbnlhbmNoYS5jb20vc2VhcmNoP2tleT1BVk5FVCUyMFRFQ0hOT0xPR1klMjBIT05HJTIwS09ORyUyMExJTUlURUQ%3D; RTYCID=9cad913b5f9648ce990f022c29afde9b; TYCID=14973500f30511e88783936a6b061dc9; _gat_gtag_UA_123487620_1=1; _gid=GA1.2.374185403.1575436487; undefined=14973500f30511e88783936a6b061dc9; CT_TYCID=655a7a562cb24f74acd14532553213c9; ssuid=7473513048; _ga=GA1.2.1024878074.1546934820; __insp_slim=1547367407173; cloud_utm=9658e8623d2148179619702cdfe13198; Hm_lvt_e92c8d65d92d534b0fc290df538b4758=1575689886,1575702980,1575703521,1575706298; tyc-user-info=%257B%2522claimEditPoint%2522%253A%25220%2522%252C%2522explainPoint%2522%253A%25220%2522%252C%2522integrity%2522%253A%25220%2525%2522%252C%2522state%2522%253A0%252C%2522announcementPoint%2522%253A%25220%2522%252C%2522bidSubscribe%2522%253A%2522-1%2522%252C%2522vipManager%2522%253A%25220%2522%252C%2522onum%2522%253A%25220%2522%252C%2522monitorUnreadCount%2522%253A%25220%2522%252C%2522discussCommendCount%2522%253A%25220%2522%252C%2522token%2522%253A%2522eyJhbGciOiJIUzUxMiJ9.eyJzdWIiOiIxMzcyMDEwODExNyIsImlhdCI6MTU3NTcwMzU2MywiZXhwIjoxNjA3MjM5NTYzfQ.1eSbT0GCR594Dt4CS1EkDgOK7wVLuBFWQbvemuxxAIlo5CyFSXj1bxY_Vvvnbn2Kawv-H4j_sS9nZ7tL-TO65w%2522%252C%2522claimPoint%2522%253A%25220%2522%252C%2522redPoint%2522%253A%25220%2522%252C%2522myAnswerCount%2522%253A%25220%2522%252C%2522myQuestionCount%2522%253A%25220%2522%252C%2522signUp%2522%253A%25221%2522%252C%2522nickname%2522%253A%2522%25E5%25A4%25A7%25E5%258D%25AB%25C2%25B7%25E6%259E%2597%25E5%25A5%2587%2522%252C%2522privateMessagePointWeb%2522%253A%25220%2522%252C%2522privateMessagePoint%2522%253A%25220%2522%252C%2522isClaim%2522%253A%25220%2522%252C%2522new%2522%253A%25221%2522%252C%2522pleaseAnswerCount%2522%253A%25220%2522%252C%2522vnum%2522%253A%25220%2522%252C%2522bizCardUnread%2522%253A%25220%2522%252C%2522mobile%2522%253A%252213720108117%2522%257D; __insp_targlpt=QVZORVQgVEVDSE5PTE9HWSBIT05HIEtPTkcgTElNSVRFRF%2Fnm7jlhbPmkJzntKLnu5Pmnpwt5aSp55y85p_l; __insp_norec_sess=true; __insp_nv=true; jsid=SEM-BAIDU-PZ1907-SY-000100; aliyungf_tc=AQAAAFwe4W9DkAEASsoRG5B5Yy5iqQGK; csrfToken=QVH_R1DRUCu4MCQvOagP2wcA',
}

def getUrlByCompanyName(name):
	'''
	input name output the url of the company
	'''
	url = "https://www.tianyancha.com/search?key=" + name

	r = requests.get(url,headers = header)
	html = etree.HTML(r.text)
	href = html.xpath('//div[@class="search-result-single   "]//a[@class="name "]/@href')

	if len(href) == 0:
		print("获取查询列表失败")
		return False
	else:
		return href[0]

def getCompanyByUrl(url,commpanyName):
	'''
	input the company url output the full html 
	'''
	company = Company()
	r = requests.get(url, headers = header)
	html = etree.HTML(r.text)
	name = html.xpath('//div[@class="content"]//h1[@class="name"]/text()')
	tel = html.xpath('//div[@class="detail "]//div[@class="f0"]/div[1]/span[2]/text()')
	email = html.xpath('//div[@class="detail "]//div[@class="f0"]/div[2]/span[2]/text()')
	orgname = commpanyName
	website = html.xpath('//div[@class="detail "]//div[@class="f0 clearfix"]/div[1]/span[2]/text()')
	adress = html.xpath('//div[@class="detail "]//div[@class="f0 clearfix"]//div[@class="auto-folder"]/div[1]/text()')
	simpleinfo = html.xpath('//div[@class="detail "]//div[@class="summary"]//text()')
	company.name = "" if len(name) == 0 else name[0]
	company.tel = "" if len(tel) == 0 else tel[0]
	company.email = "" if len(email) == 0 else email[0]
	company.website = "" if len(website) == 0 else website[0]
	company.adress = "" if len(adress) == 0 else adress[0]
	company.simpleinfo = "" if len(simpleinfo) == 0 else simpleinfo[0]
	company.orgname = orgname
	print(company)
	return company

def getCompaniesByCsv(file):
	companies = []
	with open(file, 'r', encoding= 'GBK') as csvfile:
		lines = csv.reader(csvfile)
		companies = [line[0] for line in lines]
	return companies


def writeToCsvByCompanyies(companies):
	'''
	input companies 
	'''
	with open('output.csv', 'w', encoding='utf-8', newline='') as csvfile:
		writer = csv.writer(csvfile)
		for company in companies:
			writer.writerow(company.toArray())


def main():
	companies = []
	companynames = getCompaniesByCsv('input.csv')
	try:
		for name in companynames:
			url = getUrlByCompanyName(name)
			company = getCompanyByUrl(url,name)
			companies.append(company)
	finally:
		writeToCsvByCompanyies(companies)



if __name__ == '__main__':
	main()