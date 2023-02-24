import yaml, os

if __name__ == "__main__":
	config = {
		'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/107.0.0.0 Safari/537.36', 
		'frequency_threshold': os.getenv('frequency_threshold') or 6000, 
		'wkhtmltopdf': '/usr/bin/wkhtmltopdf', 
		'msg_from': os.getenv('msg_from'), 
		'passwd': os.getenv('passwd'), 
		'msg_to': os.getenv('msg_to').split()
	}
	if not config['msg_from']:
		raise KeyError('[ERROR] Please set `msg_from` in secrets!')
	if not config['passwd']:
		raise KeyError('[ERROR] Please set `passwd` in secrets!')
	if not config['msg_to']:
		raise KeyError('[ERROR] Please set `msg_to` in secrets!')
	
	with open('config.yaml', 'w', encoding='utf-8') as f:
		yaml.dump(config, f, encoding='utf-8', allow_unicode=True)
	
	print('[INFO] `config.yaml` is successfully generated!')