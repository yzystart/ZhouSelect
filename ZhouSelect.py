import sys
from selenium import webdriver
from lxml import etree
import os
import requests
import json


config = {
    # #要操作的文件名，包括后缀名
    # "elcel_name":"",
    # #此处需要手动配置，设置起始日期
    # "yyyy-MM-dd":"2018-08-01",
    # #上一次记录完的页码
    # "page":"1"
}


# 生成配置文件
def init_config():
    config = {
        # 工作目录
        'work_directory': '',
        # 当前操作的excel
        'excel': '',
        # 上一次录入到的行
        'last_line': ''
    }
    savefile()

# 保存配置文件
def savefile():
    with open(os.path.abspath('.') + '/config.json', encoding='utf-8', mode='w') as writefile:
        writefile.write(json.dumps(config, ensure_ascii=False))

# 读取配置文件
def read_config():
    try:
        print('正在读取配置文件...')
        # 读取配置文件
        with open(os.path.abspath('.') + '/config.json', encoding='utf-8-sig', mode='r') as readfile:
            # 读取文本
            str_config = readfile.read()
            # 转化为字典
            config = json.loads(str_config)
            print(config)
    except FileNotFoundError:
        print('没有配置文件,开始生成..')
        init_config()
        input('生成结束,请重新打开程序')


def query(org):
    query_url = "https://www.tianyancha.com/search?key="+org
    # 获取分公司headers
    query_headers = {
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
        'Accept-Encoding': 'gzip, deflate, br',
        'Accept-Language': 'zh-CN,zh;q=0.9',
        'Connection': 'keep-alive',
        'Host': 'www.tianyancha.com',
        'Referer': 'https://www.tianyancha.com/',
        'sec-ch-ua': '" Not A;Brand";v="99", "Chromium";v="90", "Google Chrome";v="90"',
        'sec-ch-ua-mobile': '?0',
        'Sec-Fetch-Dest': 'document',
        'Sec-Fetch-Mode': 'navigate',
        'Sec-Fetch-Site': 'same-origin',
        'Sec-Fetch-User': '?1',
        'Upgrade-Insecure-Requests': '1',
        'Cookie': Cookie,
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/90.0.4430.72 Safari/537.36'
    }

    query_response = requests.get(url=query_url,headers=query_headers)

    html = etree.HTML(query_response.text)
    try:
        # 所有搜索列表
        resule_list_div=html.xpath("//div[@class='result-list sv-search-container']")[0]
        # 一般第一个就是正确的，取搜索列表中的第一个
        content_div=resule_list_div.xpath("./div/div/div[@class='content']/div/a/em[text()='"+org+"']/../../..")[0]
        print('***************************************')
        try:
            org_name=content_div.xpath('./../div[@class="triangle-xcx"]/div[@class="xcx-block"]/div[@class="info"]')[0].text
            # org_name=resule_list_div.xpath('./div/div[@class="search-result-single  "]/div[@class="triangle-xcx"]/div[@class="xcx-block"]/div[@class="info"]')[0].text
            print('公司名->',org_name)
        except Exception as e:
            print('没有定位到公司名',e)
        try:
            user=content_div.xpath('./div[@class="info row text-ellipsis"]/div[1]/a')[0].text
            print('法人/经营者->',user)
        except Exception as e:
            print('没有查询到法人/经营者',e)
        try:
            status= content_div.xpath("./div[@class='header']/div")[0].text
            print('经营状态 ->',status)
        except Exception as e:
            print('没有查询到经营状态',e)
        try:
            phone = content_div.xpath('.//span[text()="电话："]/../span[2]/span')[0].text
            print('电话 ->',phone)
        except Exception as e:
            print('没有查询到电话',e)
        try:
            address = content_div.xpath('.//span[text()="地址："]/../span[2]')[0].text
            print('地址 ->',address)
        except Exception as e:
            print('没有查询到地址',e)
        print('***************************************')
    except Exception as e:
        print('没有查询到该公司',e)


print('***************************************')
print('**********欢迎来到老周查询器2.0***********')
print('***************************************')
print('*************选择一个选项？**************')
print('***************************************')
print('**************1.单查询*****************')
print('**************2.Excel录入**************')
print('***************************************')

select=input()
print('>当前的操作路径为', os.path.abspath('.'))
if select=='2':
    read_config()
# 驱动路径
bro = webdriver.Chrome(executable_path=os.path.abspath('.') + '\\' + "chromedriver.exe")
# 打开网页获取cookie
bro.get("https://www.tianyancha.com/")
#
input('请在页面登录你的天眼查账号，登录完成后回到cmd窗口回车')

Cookie=''
# 设置cookie
for cookie in bro.get_cookies():
    Cookie=Cookie+cookie['name']+'='+cookie['value']+';'
Cookie=Cookie[0:-1]
print(Cookie)
if __name__ == "__main__":

    if select=='1':
        while True:
            try:
                query(input('请输入一个机构：'))
            except Exception as e:
                print(e)
                continue
    if select =='2':
        read_config()

