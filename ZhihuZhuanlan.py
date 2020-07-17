import json
import os

import requests
import pypandoc

ZhuanlanID = '1250854045965189120'

headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/83.0.4103.116 Safari/537.36'}
cookies = {
    # 浏览器复制Cookies填在此处；Cookies包含个人登录密码，请勿外传。
}

os.makedirs(ZhuanlanID)

for offset_i in range(10):
    offset = str(offset_i * 20)
    CatalogUrl = 'https://api.zhihu.com/remix/well/{}/catalog?offset={}&limit=20&order_by=global_idx'.format(ZhuanlanID, offset)
    CatalogRequ = requests.get(url = CatalogUrl, headers = headers, cookies = cookies)

    coo = CatalogRequ.cookies
    CooRequ = requests.get(url = CatalogUrl, headers = headers)

    CatalogDict = json.loads(CatalogRequ.text)
    if not CatalogDict['data']:
        break
    for SectionInfo in CatalogDict['data']:
        SectionIndex = SectionInfo['index']['serial_number_txt']
        SectionTitle = SectionInfo['title']
        SectionID = SectionInfo['id']
        SectionUrl = 'https://www.zhihu.com/market/paid_column/{}/section/{}'.format(ZhuanlanID, SectionID)
        SectionRequ = requests.get(url = SectionUrl, headers = headers, cookies = cookies)
        with open('temp.html', 'w+', encoding='utf-8') as html:
            html.write(SectionRequ.text)
        pypandoc.convert_file('temp.html', 'docx', outputfile = '{}/{} {}.docx'.format(ZhuanlanID, SectionIndex, SectionTitle))
