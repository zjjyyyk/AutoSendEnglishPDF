import yaml
import requests
from lxml import etree
import pdfkit
import pandas as pd
import re
import random
import os
import time

# 邮件相关
import smtplib
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
import pyjokes


class YamlLoader:
	def __init__(self,yaml_filename='config.yaml'):
		self.yaml_list = []
		with open(yaml_filename,'rb') as f:
			self.yaml_list = list(yaml.safe_load_all(f))

	def get_params_dict(self,list_index:int):
		return self.yaml_list[list_index]


class NewsCrawler:
	def __init__(self,user_agent):
		self.headers = {
			'User-Agent': user_agent
		}
		self.forbidden_list = []

	## private
	def _get_latest_i_news_url(self,i:int):
		url = 'https://bbcworldnews.net/technology/'
		response = requests.get(url,headers=self.headers)
		selector = etree.HTML(response.text)
		newsURL = selector.xpath('//*[@id="content_masonry"]/div[{}]/div/div[2]/div/div/h3/a/@href'.format(i))[0]
		return newsURL

	## public ##
	def load_forbiddens_from_file(self,filename:str):
		with open(filename,'r') as f:
			self.forbidden_list = f.readlines()

	def add_forbidden_url(self,url:str):
		self.forbidden_list.append(url)
	
	def write_forbiddens_to_file(self,filename:str):
		with open(filename,'w') as f: 
			for url in self.forbidden_list:
				f.write(url+'\n')

	def get_valid_news_url(self):
		try:
			i = 1
			newsURL = self._get_latest_i_news_url(i)
			while newsURL in self.forbidden_list:
				i = i + 1
				newsURL = self._get_latest_i_news_url(i)
			return newsURL
		except:
			print("Error happened when i =",i)
			raise

	def get_news_response(self,url:str):
		response = requests.get(url,headers=self.headers)
		return response


class BBC_ResponseParser:
	def __init__(self,response):
		self.response = response
		self.content = self._extract_response_content_as_dict(response)
		self.wordList = set()
		self.is_parse_difficult_words = True
		self.parse_difficult_words_params = None
		self.difficult_words = set()
		self.html = ''

	## private
	@staticmethod
	def _extract_response_content_as_dict(response):
		selector = etree.HTML(response.text)  # html为Element对象
		title = selector.xpath('/html/body/div[1]/div/section/div/div/div[1]/div/div[1]/div/div[1]/div[2]/h1/text()')[0].strip()
		datetime = selector.xpath('/html/body/div[1]/div/section/div/div/div[1]/div/div[1]/div/div[1]/div[2]/span[2]/span/@datetime')[0]
		p_elements = selector.xpath('/html/body/div[1]/div/section/div/div/div[1]/div/div[1]/div/div[2]/p') + \
					selector.xpath('/html/body/div[1]/div/section/div/div/div[1]/div/div[1]/div/div[2]/div/p')
		paragraphs = [''.join(p.xpath('.//text()')) for p in p_elements]
		content = {
			'title':title,
			'datetime':datetime,
			'paragraphs':paragraphs
		}
		return content

	def _clean_and_wrap_raw_content_while_update_wordlist(self,sentence:str, type:str):
		assert type in ['title', 'datetime', 'paragraph'], 'unvalid type in "clean_and_wrap_data"'
		sentence = sentence.replace("’","'").replace('“','"').replace('”','"').replace('—','--').strip()
		self.wordList.update([''.join(filter(str.isalpha,word.lower())) for word in re.split(r'[ ,.?!-a():;]',sentence) if word != ''])
		wrap_dict = {
				'title': '<h1 class="title">{}</h1>',
				'datetime': '<div class="time">{}</div>',
				'paragraph': '<em class="paragraph">{}</em>'
		}
		html_block = wrap_dict[type].format(sentence)
		return html_block

	@staticmethod
	def _choose_difficult_words(wordList:set,frequency_xlsx_path:str,sheet_name:str,frequency_threshold:int,num:int):
		assert len(wordList) != 0 , '无法对空WordList进行选择操作'
		df = pd.read_excel(frequency_xlsx_path,sheet_name=sheet_name,engine='openpyxl')
		difficult_words = {word.lower() for word in wordList if '  '+word.lower() in df[' word'].values and df[df[' word']=='  '+word.lower()]['RANK #'].values[0] > frequency_threshold}
		while len(difficult_words) < num:
			frequency_threshold = frequency_threshold - 100
			difficult_words = {word.lower() for word in wordList if '  '+word.lower() in df[' word'].values and df[df[' word']=='  '+word.lower()]['RANK #'].values[0] > frequency_threshold}
		print('Current frequency_threshold:',frequency_threshold)
		return frequency_threshold, random.sample(difficult_words,num)

	## public ##
	def not_parse_difficult_words(self):
		self.is_parse_difficult_words = False

	def parse_difficult_words(self,frequency_xlsx_path:str,sheet_name:str,frequency_threshold:int,num:int = 4):
		self.is_parse_difficult_words = True
		self.parse_difficult_words_params = (frequency_xlsx_path,sheet_name,frequency_threshold,num)

	def format_html(self,save_path:str=None):
		html_title = self._clean_and_wrap_raw_content_and_update_wordlist(self.content['title'],'title')
		html_datetime = self._clean_and_wrap_raw_content_and_update_wordlist(self.content['datetime'],'datetime')
		html_paragraphs = [self._clean_and_wrap_raw_content_and_update_wordlist(paragraph,'paragraph') for paragraph in self.content['paragraphs'] ]
		html_body = '\n'.join(html_paragraphs)
		html_random_words = ''
		if self.is_parse_difficult_words:
			frequency_threshold, random_words = self._choose_difficult_words(self.wordList,self.parse_difficult_words_params)
			print(random_words)
			html_random_words = '<strong class="random_words">{}</strong>'.format('Selected words:&emsp;'+'&emsp;'.join(random_words))
		foreHtml = '''
			<!DOCTYPE html>
			<html>
			<head>
			    <meta name="viewport" content="width=device-width, initial-scale=1.0">
				<meta charset="utf-8">
			    <title>Document</title>
			    <style type="text/css">
			        h1.title {
			            -webkit-text-size-adjust: 100%;
			            -webkit-tap-highlight-color: rgba(0,0,0,0);
			            box-sizing: border-box;
			            margin: .67em 0;
			            letter-spacing: -.03em;
			            font-weight: 600 !important;
			            text-transform: capitalize !important;
			            font-family: Poppins !important;
			            font-size: 40px;
			            color: #000;
			            margin-top: 18px;
			            margin-bottom: 0;
			            line-height: 1.125!important;
			            float: left;
			            width: 100%;
			            text-align: left;
			        }
			        .time {
			            -webkit-text-size-adjust: 100%;
			            -webkit-tap-highlight-color: rgba(0,0,0,0);
			            line-height: 1.85;
			            text-align: left;
			            box-sizing: border-box;
			            font-size: 12px;
			            margin: 0 5px;
			            padding-right: 5px;
			            text-transform: capitalize;
			            letter-spacing: .03em;
			            font-weight: 400;
			            font-family: Poppins !important;
			            color: #000!important;
			        }
			        em.paragraph {
			            display: block;
			            -webkit-text-size-adjust: 100%;
			            -webkit-tap-highlight-color: rgba(0,0,0,0);
			            line-height: 1.55;
			            margin-bottom:12px;
			            color: #666;
			            font-family: Poppins !important;
			            font-weight: 400 !important;
			            font-size: 20px;
			            box-sizing: border-box;
			        }
					strong.random_words {
						    -webkit-text-size-adjust: 100%;
							-webkit-tap-highlight-color: rgba(0,0,0,0);
							list-style: none;
							line-height: 1.3;
							font-family: Poppins !important;
							box-sizing: border-box;
							background-color: transparent;
							letter-spacing: -.03em;
							margin-top: calc(13px + 1.2em);
							margin-bottom: calc(10px + .2em);
							font-weight: 600 !important;
							text-transform: capitalize !important;
							outline: 0;
							transition: all 0.4s ease 0s;
							text-decoration: none!important;
							color: #666666 !important;
							font-size: 17px;
					}
			    </style>
			</head>
			<body>
			<div> <img src="https://bbcworldnews.net/wp-content/uploads/2021/01/bbc-3.png" height=50> </div>
					'''
		backHtml = '''
			{title}
			{datetime}
			{body}
			{random_words}
			</body>
			</html>
		'''.format(title=html_title, datetime = html_datetime, body = html_body, random_words = html_random_words)
		html = foreHtml + backHtml
		if save_path:
			with open(save_path,'w') as f:
				f.write(html)
		self.html = html
		return html

	def format_pdf(self,wkhtmltopdf_path):
		if self.html == '':
			self.format_html('temp.html')
		title = self.content.get('title','NoTitle')
		config = pdfkit.configuration(wkhtmltopdf=wkhtmltopdf_path)
		pdfkit.from_string(self.html,'{}.pdf'.format(title),configuration=config)
		print("转换完成,输出{}.pdf".format(title))
		return '{}.pdf'.format(title)


class MailSender:
	def __init__(self,msg_from,passwd):
		self.msg_from = msg_from
		self.passwd = passwd

	## private ##

	## public
	@staticmethod
	def send_email(subject,msg_from, passwd, msg_to, text_content, file_path=None,max_chances=10):
		print('开始发送啦')
		msg = MIMEMultipart()
		subject = "每日英语新闻阅读"  # 主题
		text = MIMEText(text_content)
		msg.attach(text)

		# file_path = r'read.md'  #如果需要添加附件，就给定路径
		if file_path:  # 最开始的函数参数我默认设置了None ，想添加附件，自行更改一下就好
			docFile = file_path
			docApart = MIMEApplication(open(docFile, 'rb').read())
			docApart.add_header('Content-Disposition', 'attachment', filename=docFile)
			msg.attach(docApart)
			print('发送附件！')
		msg['Subject'] = subject
		msg['From'] = msg_from
		msg['To'] = msg_to
		for i in range(max_chances):
			try:
				s = smtplib.SMTP_SSL("smtp.qq.com", 465)
				s.login(msg_from, passwd)
				s.sendmail(msg_from, msg_to, msg.as_string())
				print("发送成功",msg_to)
				s.quit()
				return True
			except smtplib.SMTPException as e:
				print("发送失败",msg_to,'剩余尝试次数：',max_chances-i-1)
				time.sleep(5)
				continue
		print("所有尝试均失败了，无奈返回False")
		s.quit()
		return False 

	def send_multi_email(self, subject,msg_to_list:list,text_content, file_path=None,max_chances=10):
		for msg_to in msg_to_list:
			if not self.send_email(subject,self.msg_from, self.passwd, msg_to, text_content, file_path,max_chances):
				print('发送邮件时出现问题')
				return False
		return True

class CodeAdmin:
	def __init__(self):
		os.chdir(os.path.dirname(os.path.abspath(__file__)))
		# MORE PERSONALIZED SETTING
		pass

	## private

	## public
	def register_logInfo(self,info:str):
		pass

	def run_code(main_func):
		try:
			print('Start time:',time.strftime('%Y-%m-%d %H:%M:%S',time.localtime(time.time())))
			main_func()
			print('End time:',time.strftime('%Y-%m-%d %H:%M:%S',time.localtime(time.time())))
		except Exception as e:
			pass
			raise e

def main():
	params = YamlLoader(yaml_filename='config.yaml').get_params_dict(0)
	crawler = NewsCrawler(user_agent=params['User-Agent'])
	crawler.load_forbiddens_from_file(filename='history.txt')
	valid_news_url = crawler.get_valid_news_url()
	response = crawler.get_news_response(url=valid_news_url)
	parser = BBC_ResponseParser(response=response)
	parser.parse_difficult_words(frequency_xlsx_path='word frequency list 60000 English.xlsx',sheet_name='Sheet1', 
		frequency_threshold=params['frequency_threshold'],num=4)
	pdf_filename = parser.format_pdf(wkhtmltopdf=params['wkhtmltopdf'])
	sender = MailSender(msg_from=params['msg_from'],passwd=params['passwd'])
	isAllSuccess = sender.send_multi_email(
			subject = '每日英语新闻阅读',
			msg_to_list = params['msg_to'],
			text_content = "Hi, today!\n\n" + pyjokes.get_joke(),
			file_path = pdf_filename,
			max_chances = 10
		)
	if isAllSuccess:
		crawler.add_forbidden_url(valid_news_url)
		crawler.write_forbiddens_to_file(filename='history.txt')
		os.remove(pdf_filename)

if __name__ == '__main__':
	zjj = CodeAdmin()
	zjj.run_code(main)



