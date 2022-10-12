import yaml
import httpx as requests
from lxml import etree
import pdfkit

# 邮件相关
import smtplib
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
import pyjokes

'''
TODO
0. requirements.txt
1. 从https://wkhtmltopdf.org/downloads.html下载安装对应版本，并修改下面configuration里的路径
2. 修改headers？
3. 还可以调一调style
'''

# 读取配置文件
params = {}
with open('config.yaml','rb') as f:
	params = list(yaml.safe_load_all(f))[0]


headers = {
	'User-Agent': params['User-Agent']
}

def get_latest_news_url(i:int):
	url = 'https://bbcworldnews.net/technology/'
	response = requests.get(url,headers=headers)
	selector = etree.HTML(response.text)  # html为Element对象
	newsURL = selector.xpath('//*[@id="content_masonry"]/div[{}]/div/div[2]/div/div/h3/a/@href'.format(i))[0]
	return newsURL

def get_news_content(url:str):
	response = requests.get(url,headers=headers)
	selector = etree.HTML(response.text)  # html为Element对象
	title = selector.xpath('/html/body/div[1]/div/section/div/div/div[1]/div/div[1]/div/div[1]/div[2]/h1/text()')[0].strip()
	datetime = selector.xpath('/html/body/div[1]/div/section/div/div/div[1]/div/div[1]/div/div[1]/div[2]/span[2]/span/@datetime')[0]
	paragraphs = selector.xpath('/html/body/div[1]/div/section/div/div/div[1]/div/div[1]/div/div[2]//text()')
	dummy_linkContents = selector.xpath('/html/body/div[1]/div/section/div/div/div[1]/div/div[1]/div/div[2]//a/text()')
	for linkContent in dummy_linkContents:
		index = paragraphs.index(linkContent)
		if index == 0:
			paragraphs[index + 1] = paragraphs[index] + paragraphs[index + 1]
			del paragraphs[index]
		elif index == len(paragraphs)-1:
			paragraphs[index - 1] = paragraphs[index - 1] + paragraphs[index]
			del paragraphs[index]
		else:
			paragraphs[index - 1] = paragraphs[index - 1] + paragraphs[index] + paragraphs[index + 1]
			del paragraphs[index]
			del paragraphs[index]
	content = {
		'title':title,
		'datetime':datetime,
		'paragraphs':paragraphs
	}
	return content

def clean_and_wrap_data(sentence:str, type:str):
	assert type in ['title', 'datetime', 'paragraph'], 'unvalid type in "clean_and_wrap_data"'
	sentence = sentence.replace("’","'").replace('“','"').replace('”','"').strip()
	wrap_dict = {
			'title': '<h1 class="title">{}</h1>',
			'datetime': '<div class="time">{}</div>',
			'paragraph': '<em class="paragraph">{}</em>'
	}
	html_block = wrap_dict[type].format(sentence)
	return html_block



def format_pdf(content:dict):
	html_title = clean_and_wrap_data(content['title'],'title')
	html_datetime = clean_and_wrap_data(content['datetime'],'datetime')
	html_paragraphs = [clean_and_wrap_data(paragraph,'paragraph') for paragraph in content['paragraphs'] ]
	html_body = '\n'.join(html_paragraphs)
	foreHtml = '''
<!DOCTYPE html>
<html>
<head>
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
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
            margin-top: 0;
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
    </style>
</head>
<body>
		'''
	backHtml = '''
{title}
{datetime}
{body}
</body>
</html>

	'''.format(title=html_title, datetime = html_datetime, body = html_body)
	html = foreHtml + backHtml
	try:
		config = pdfkit.configuration(wkhtmltopdf=params['wkhtmltopdf'])
		pdfkit.from_string(html,'{}.pdf'.format(content['title']),configuration=config)
		print("转换完成,输出{}.pdf".format(content['title']))
		return '{}.pdf'.format(content['title'])
	except:
		print('ERROR occured.')
		raise

def send_to_qqMail(pdf_filename):
	def send_email(msg_from, passwd, msg_to, text_content, file_path=None):
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
		try:
		  s = smtplib.SMTP_SSL("smtp.qq.com", 465)
		  s.login(msg_from, passwd)
		  s.sendmail(msg_from, msg_to, msg.as_string())
		  print("发送成功")
		  s.quit()
		  return True
		except smtplib.SMTPException as e:
		  print("发送失败")
		  s.quit()
		  return False
	msg_from = params['msg_from']  # 发送方邮箱
	passwd = params['passwd']  # 填入发送方邮箱的授权码（就是刚刚你拿到的那个授权码）
	msg_to_list = params['msg_to']  # 收件人邮箱，我是自己发给自己
	text_content = "Hi, today!\n\n" + pyjokes.get_joke() # 发送的邮件内容
	file_path = pdf_filename # 需要发送的附件目录
	for msg_to in msg_to_list:
		if not send_email(msg_from,passwd,msg_to,text_content,file_path):
			print('发送邮件时出现问题')
			break



if __name__ == '__main__':
	with open('history.txt','a+') as f:
		i = 1
		news_url = get_latest_news_url(i)
		while news_url in f.readlines():
			i = i + 1
			news_url = get_latest_news_url(i)
		print('i =',i,'news_url:',news_url)
		content = get_news_content(news_url)
		pdf_filename = format_pdf(content)
		isSuccess = send_to_qqMail(pdf_filename)
		if isSuccess:
			pass
			# f.write(news_url+'\n')
		else:
			print('Something wrong happened.')


		









