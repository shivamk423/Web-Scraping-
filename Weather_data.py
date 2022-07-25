from cgitb import text
from email import header
from requests_html import HTMLSession
session = HTMLSession()
query = 'mumbai'
url = f'https://www.google.com/search?q=weather+{query}'
r = session.get(url, headers ={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/103.0.0.0 Safari/537.36'})
temp = r.html.find('span#wob_tm', first=True).text
unit = r.html.find('div.vk_bk.wob-unit span.wob_t',first=True).text
desc= r.html.find('div.VQF4g',first=True).find('span#wob_dc',first=True).text
print(query,temp,unit,desc)