# AutoSendEnglishPDF

> 程序很简单，学英语是关键

## one-sentence usage
每天从BBC爬一篇新闻下来转成pdf发到qq邮箱里，便于打印学习

## 使用方法
1. 安装依赖 `pip install -r requirements.txt`
2. 去 https://wkhtmltopdf.org/downloads.html 下载对应版本的wkhtmltopdf并安装
3. 创建一个`config.yaml`，内容形式是
```
User-Agent: Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/107.0.0.0 Safari/537.36  # 设置headers

wkhtmltopdf: G:\wkhtmltopdf\bin\wkhtmltopdf.exe  # 设置wkhtmltopdf运行路径

msg_from: 791xxxxxxx9@qq.com  # 设置发送者的邮箱
passwd: xxxxxxxxxxx  # smtp授权码（https://blog.csdn.net/weixin_62650212/article/details/123829431）
msg_to:  # 设置要发送给谁的邮箱，支持1个或多个
  - 791xxxxxxx9@qq.com  
  - xxx@qq.com
```
4. 购买一个centos云服务器，然后`crontab -e`接`0 7 * * * python [DIRPATH]/BBC_crawl.py` （每天7点执行一次） ， 或者issue你的邮箱，我把你加进msg_to，大家一起学习:sunny: 
