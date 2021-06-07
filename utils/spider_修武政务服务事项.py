# coding:utf-8
import requests
import requests
import os,random,re
from lxml import etree
import json
import pandas as pd
import math
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
import time
import pandas as pd
from lxml import etree
from datetime import datetime
from pathlib import Path

class Httprequest(object):
    ua_list = [
    'Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/535.1 (KHTML, like Gecko) Chrome/14.0.835.163 Safari/535.1','Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/73.0.3683.103 Safari/537.36Chrome 17.0',
    'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_7_0) AppleWebKit/535.11 (KHTML, like Gecko) Chrome/17.0.963.56 Safari/535.11',
    'Mozilla/5.0 (Windows NT 6.1; WOW64; rv:6.0) Gecko/20100101 Firefox/6.0Firefox 4.0.1',
    'Mozilla/5.0 (Macintosh; Intel Mac OS X 10.6; rv:2.0.1) Gecko/20100101 Firefox/4.0.1',
    'Mozilla/5.0 (Macintosh; U; Intel Mac OS X 10_6_8; en-us) AppleWebKit/534.50 (KHTML, like Gecko) Version/5.1 Safari/534.50',
    'Mozilla/5.0 (Windows; U; Windows NT 6.1; en-us) AppleWebKit/534.50 (KHTML, like Gecko) Version/5.1 Safari/534.50',
    'Opera/9.80 (Windows NT 6.1; U; en) Presto/2.8.131 Version/11.11',
    ]
    #把方法变成属性的装饰器
    @property  
    def random_headers(self):
        return {
            'User-Agent': random.choice(self.ua_list)
        }


class ProjectContent(object):
    
    def __init__(self,pagenum,cityid,deptname,totalnum,jsessionid):
        self.url = "http://59.207.104.2:8060/smp/app/module/default/jsp/view/view.action"
        self.headers ={
                    'Connection': 'keep-alive',
                    'Content-Length': '2214',
                    'Host': '59.207.104.2:8060 ' , 
                    'Origin':'http://59.207.104.2:8060',
                    'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
                    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/84.0.4147.125 Safari/537.36',
                    'X-Requested-With': 'XMLHttpRequest',
                    'Accept': 'application/json, text/javascript, */*; q=0.01',
                    'Accept-Encoding': 'gzip, deflate',
                    'Accept-Language': 'zh-CN,zh;q=0.9', 
                    'Cookie': f'JSESSIONID={jsessionid}',  
                    }
        self.pagenum=pagenum
        self.totalnum=totalnum
        self.cityid= cityid
        self.deptname=deptname

    
    
    def get_projectlist(self):
        payload = {
                "fn": "grid_list",
                "viewId": "2E43A79F6428539CDCB3501D91AED5D1",
                "sysUnid": 'D0971696C2081F86689F1FE0851AA4C6',
                "ID": f'{self.cityid}003',
                "DEPTUNID": f'{self.cityid}003',
                "DEPTNAME": f'{self.deptname}',
                "PARENTUNID": '0',
                "SERVICE_NAME": '',
                "TASKHANDLEITEM": '',
                "SERVICE_CODE": '',
                "page": f"{self.pagenum}",
                "rows": '100',
                "total": f"{self.totalnum}",
                }
        requests.packages.urllib3.disable_warnings()
        resp = requests.post(self.url, data=payload, headers=self.headers,verify=False)
        pro_contents=resp.json()
        if 'rows' in pro_contents.keys():
            pro_lists=pro_contents['rows']
            if len(pro_lists) > 0:
                for pro_list in pro_lists:
                    dept_name=pro_list["DEPT_NAME"]
                    servicename=pro_list["SERVICE_NAME"]
                    servicecode=pro_list["SERVICE_CODE"]
                    unid=pro_list["UNID"]
                    dept_unid=pro_list["DEPTUNID"]
                    promisedays=pro_list["PROMISEDAYS"],
                    legaldays=pro_list["LEGALDAYS"]
                    flag=pro_list["FLAG"]
                    content={
                        '办理单位':dept_name,
                        "业务办理项名称":servicename,
                        "事项编码":servicecode,
                        "unid":unid,
                        "dept_unid":dept_unid,
                        "承诺时限":promisedays,
                        "法定时限":legaldays,
                        "事项绑定":flag,
                    }
                    print(content)
                    contents.append(content)
        write_excle(contents,savefile)
        

class Test_loggin(object):
 
    def __init__(self, login_url,cookie_path='./source',cookie_name = 'cookies_sxglk.txt', expiration_time = 30):
        '''
        :param login_url: 登录网址
        :param home_url: 首页网址
        :param cookie_path: cookie文件存放路径
        :param cookie_path: 文件命名
        :param expiration_time: cookie过期时间,默认30分钟
        '''
        self.login_url = login_url
        self.cookie_path = cookie_path
        self.cookie_name = cookie_name
        self.expiration_time = expiration_time
 

    def get_cookie(self):
        '''登录获取cookie'''
        '''登录获取cookie'''
        #设置driver启动参数
        driver_path='.\source\chromedriver.exe'
        option = Options()
        option.add_experimental_option('excludeSwitches', ['enable-automation'])
        option.add_experimental_option('useAutomationExtension', False)
        driver = webdriver.Chrome(executable_path=driver_path,options=option)
        driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
            "source": """
                Object.defineProperty(navigator, 'webdriver', {
                get: () => undefined
                })
            """
            })
        driver.maximize_window()
        driver.implicitly_wait(10)
        #完成获取cookie步骤
        driver.get(self.login_url)
        driver.find_element_by_name('username').send_keys(USENAME)
        driver.find_element_by_name('password').send_keys(PASSWORD)
        driver.find_element_by_name('noLogin').click()
        driver.find_element_by_id('btn-submit-login').click()
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CLASS_NAME, 'system')))
        driver.find_element_by_xpath(r'//li[@class="appLi"][1]').click()
        driver.switch_to.window(driver.window_handles[-1])
        time.sleep(5)
        with open('./source/cookies_sxglk.txt', 'w') as cookief:
        # 将cookies保存为json格式
            cookief.write(json.dumps(driver.get_cookies()))
        print("cookies保存完成")
        driver.close()

    
    def judge_cookie(self):
        '''获取最新的cookie文件，判断是否过期'''
        my_file = Path("./source/cookies_sxglk.txt")
        if my_file.is_file():
            new_cookie = os.path.join(self.cookie_path, "cookies_sxglk.txt")
            #new_cookie = os.path.join(self.cookie_path, cookie_list2[-1])    # 获取最新cookie文件的全路径 
            file_time = os.path.getmtime(new_cookie)  # 获取最新文件的修改时间，返回为时间戳1590113596.768411
            t = datetime.fromtimestamp(file_time)  # 时间戳转化为字符串日期时间
            print('当前时间：', datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
            print('最新cookie文件修改时间：', t.strftime("%Y-%m-%d %H:%M:%S"))
            date = (datetime.now() - t).seconds // 60  # 时间之差，seconds返回相距秒数//60,返回分数
            print('相距分钟:{0}分钟'.format(date))
            if date > self.expiration_time:  # 默认判断大于30分钟，即重新手动登录获取cookie
                print("cookie已经过期，请重新登录获取")
                return self.get_cookie()
            else:
                print("cookie未过期")
        else:
            self.get_cookie()


    def get_jsessionid(self):
        '''获取JSESSIONID操作'''
        self.judge_cookie()  # 首先判断cookie是否已获取，是否过期
        print("获取JSESSIONID中")
        with open(os.path.join(self.cookie_path, self.cookie_name),'r') as cookief:
            #使用json读取cookies 注意读取的是文件 所以用load而不是loads
            cookieslist = json.load(cookief)
            # 方法1删除该字段
            cookies_dict = dict()
            for cookie in cookieslist:
                #该字段有问题所以删除就可以,浏览器打开后记得刷新页面 有的网页注入cookie后仍需要刷新一下
                if 'expiry' in cookie:
                    del cookie['expiry']
                cookies_dict[cookie['name']] = cookie['value']
        print(cookies_dict)
        jsessionid=cookies_dict['JSESSIONID']
        return jsessionid


def write_excle(content,savefile):
    df=pd.DataFrame.from_dict(content)
    df.set_index(df.columns[0],inplace=True)
    df.to_excel(savefile)


if __name__ == '__main__':
    cityid='001003018006016'
    cityname='修武县'
    deptname="修武县县直机构"
    contents = []
    #设置文件保存的地址
    savefile = ".\date\{}政务服务发布库链接.xlsx".format(cityname)
    login_url = 'http://59.207.104.12:8090//login'
    USENAME="焦作市修武县政数局_政务2019"
    PASSWORD="Abc123#$"
    totalnum=1770
    pagestart=1
    pageend=18
    test_loggin = Test_loggin(login_url=login_url)
    jsessionid=test_loggin.get_jsessionid()
    for i in range(pagestart,pageend+1):
        procontens=ProjectContent(i,cityid,deptname,totalnum,jsessionid).get_projectlist()
    write_excle(contents,savefile)
    print("{}保存完毕".format(savefile))
