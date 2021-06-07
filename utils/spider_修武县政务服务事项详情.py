# -*- coding:utf-8 -*-
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.chrome.options import Options
import time
import json
import re
import requests
from lxml import etree
import pandas as pd


class Basepage(object):
    # 构造函数
    def __init__(self, driver):
        self.driver = driver

    # 定位元素
    def locator(self, loc):
        return self.driver.find_element(*loc)

    # 关闭浏览器
    def quit(self):
        time.sleep(2)
        self.driver.quit()

    # 访问url
    def visit(self, url):
        self.driver.get(url)

    # 定位元素，输入文字
    def input_text(self, txt, loc):
        self.locator(loc).send_keys(txt)

    #定位元素，清除元素原有文字
    def clear(self,loc):
        self.locator(loc).clear()

    # 定位元素，鼠标点击
    def chick(self, loc):
        self.locator(loc).click()


class LoginPage(Basepage):
    #登录页面需要传入的参数
    loginurl='http://59.207.104.12:8090//login'
    username="焦作市修武县政数局_政务2019"
    password="Abc123#$"
    #构造函数
    def __init__(self,driver):
        self.driver = driver
        self.username_input = (By.ID,'login-by')
        self.pwd_input = (By.ID,'password')
        self.login_btn = (By.ID,'btn-submit-login')


    def to_login(self):
        self.visit(self.loginurl)
        self.input_text(self.username,self.username_input)
        self.input_text(self.password,self.pwd_input)
        self.chick(self.login_btn)
        return  DatebasePage(self.driver)


class DatebasePage(Basepage):
    
    
    def __init__(self,driver):
        self.driver = driver
        self.index_ele = (By.CLASS_NAME, 'system')
        self.index_close = (By.CLASS_NAME, 'close_2')
        self.index_url= "http://59.207.104.2:8060/smp/app/module/default/jsp/view/ztree_view.jsp?viewId=2E43A79F6428539CDCB3501D91AED5D1"
        self.new_btn = (By.ID, 'tree_2_span')
        self.input_num=(By.XPATH,r'//div[@class="datagridF"]/div/div/div/input')
        self.total_btn=(By.ID,"queryTotalButton")


    def get_datebasepage(self):
        WebDriverWait(self.driver, 10).until(EC.presence_of_element_located(self.index_ele))
        self.chick(self.index_ele)
        time.sleep(2) 
        self.driver.switch_to.window(self.driver.window_handles[-1])      
        with open('./date/cookies_bjxxk.txt', 'w') as cookief:
        # 将cookies保存为json格式
            cookief.write(json.dumps(self.driver.get_cookies()))
        self.chick(self.index_close)
        self.driver.switch_to.window(self.driver.window_handles[-1])
        self.visit(self.index_url)
        time.sleep(2)


class HandleExcle(object):
    
    def __init__(self,filename):
        self.filename = filename
        self.df=pd.read_excel(self.filename)
        self.base_url="http://59.207.104.2:8060/smp/asmp/jsp/service/service_edit.jsp?unid={}&parentunid=undefined&deptunid={}&savelogo=1&dialogId=2E43A79F6428539CDCB3501D91AED5D1"


    def get_unid(self,x):
        unid=self.df['unid'].at[x]
        return unid

  
    def get_deptunid(self,x):
        deptunid=self.df['dept_unid'].at[x]
        return deptunid
    

    def get_url(self):
        for i in range(0,self.df.shape[0]):
            unid=self.get_unid(i)
            deptunid=self.get_deptunid(i)
            url=self.base_url.format(unid,deptunid)
            parse_url(url)
            if i%100==0:
                write_excle(contents,savefile)


def read_cookies():
    cookies_dict = dict()
    with open("./date/cookies_bjxxk.txt", "r") as fp:
        cookies = json.load(fp)
        for cookie in cookies:
            cookies_dict[cookie['name']] = cookie['value']
    return cookies_dict


def four_standard_name(text):
    result_text=re.findall(r'id="four_standard".*?checked.*?>(.*?)<',text,re.U)
    four_standard='/'.join(result_text)
    return four_standard


def onlineHandleDepth_name(text):
    result_text=re.findall(r'id="onlineHandleDepth".*?checked.*?>(.*?)<',text,re.U)
    onlineHandleDepth='/'.join(result_text)
    return onlineHandleDepth


def enterprise_zt_name(text):
    content=for_user_name(text)
    if content=='自然人':
        return None
    else:
        result_text=re.findall(r'id="enterprise_zt".*?checked.*?>(.*?)<',text,re.U)
        enterprise_zt='/'.join(result_text)
        return enterprise_zt


def person_zt_name(text):
    result_text=re.findall(r'id="person_zt".*?checked.*?>(.*?)<',text,re.U)
    person_zt='/'.join(result_text)
    return person_zt


def for_user_name(text):
    result_text=re.findall(r'id="for_user".*?checked.*?>(.*?)<',text,re.U)
    for_user='/'.join(result_text)
    return for_user


def asa_unid(text):
    result_text=re.findall(r'id="asa_unid".*?checked.*?>(.*?)</tr>',text,re.U)[0]
    asaunid_name=re.findall(r'input.*?value.*?>(.*?)<',result_text,re.U)[0]
    day_name=re.findall(r'input_num.*?value="(.*?)"',result_text,re.U)[0]
    asmaxinfo_name=re.findall(r'asa_maxinfo.*?value="(.*?)"',result_text,re.U)[0]
    return asaunid_name,day_name,asmaxinfo_name


def isonline_name(text):
    isonline = re.findall(r"id=\"isOnline\".*?checked>(.*?)<", text, re.U)[-1]
    return isonline


def web_apply_url_name(text):
    content=isonline_name(text)
    if content == '是':
       web_apply_url=re.findall(r'id=\"web_apply_url\".*?value=\"(.*?)\"',text,re.U)[0]
       return web_apply_url
    else:
       return None


def webapply_name(text):
    content=isonline_name(text)
    if content == '是':
       webapplyname=re.findall(r'id=\"webapplyname\".*?value=\"(.*?)\"',text,re.U)[0]
       return webapplyname
    else:
       return None


def isEntryCenter_name(text):
    isEntryCenter=re.findall(r"id=\"isEntryCenter\".*?checked>(.*?)<", text, re.U)[-1]
    return isEntryCenter


def implement_code(text):
    implement_codenum=re.findall(r'id="implement_code".*?value=\"(.*?)\"',text,re.U)[0]
    return implement_codenum


def get_service_code(text):
    service_code = re.findall(r'id="service_code".*?value=\"(.*?)\"',text,re.U)[0]
    return service_code

def run_system_name(text):
    content = isonline_name(text)
    if content == '是':
        try:
            run_system = re.findall(r"id=\"run_system\".*?checked>(.*?)<", text, re.U)[-1]
        except IndexError:
            run_system=''
        return run_system
    else:
        return None


def service_type_name(text):
    service_type = re.findall(r"id=\"service_type\".*?checked.*?>(.*?)<", text, re.U)[-1]
    return service_type


def parse_url(url):
    cookies=read_cookies()
    resp=requests.post(url,headers=headers,cookies=cookies)
    if resp.status_code==200:
        with open(r'.\date\1.html','w',encoding='utf-8') as f:
            f.write(resp.text)
        html=etree.HTML(resp.text)
        dept_name=html.xpath('//input[@id="dept_name"]/@value')[0]
        star_level=html.xpath('//input[@id="star_level"]/@value')[0]
        catalog_name=html.xpath('//input[@id="catalog_name"]/@value')[0]
        service_name=html.xpath('//input[@id="service_name"]/@value')[0]
        promise_days=html.xpath('//input[@id="promise_days"]/@value')[0]
        legal_days=html.xpath('//input[@id="legal_days"]/@value')[0]
        web_apply_url=web_apply_url_name(resp.text)
        webapplyname=webapply_name(resp.text)
        try:
            apply_condition_desc=html.xpath('//textarea[@id="apply_condition_desc"]/text()')[0]
        except IndexError:
            apply_condition_desc=''
        try:
            legal_standard = html.xpath('//textarea[@id="legal_standard"]/text()')[0]
        except IndexError:
            legal_standard=''
        try:
            flow_desc=html.xpath('//textarea[@id="flow_desc"]/text()')[0]
        except IndexError:
            flow_desc=''
        try:
            powerflowimg=html.xpath('//img[@id="powerFlowImg"]/@src')[0]
        except IndexError:
            powerflowimg='' 
        try:
            mediationservices=html.xpath('//textarea[@id="mediation_services"]/text()')[0]
        except IndexError:
            mediationservices='' 
        try:
            result_name=html.xpath('//input[@id="result_name"]/@value')[0]
        except IndexError:
             result_name=''
        try:
            result_desc=html.xpath('//input[@id="result_desc"]/@value')[0]
        except IndexError:
            result_desc=''
        try:
            resultexampletable=html.xpath('//table[@id="resultExampleTable"]//td/a/text()')[0]
        except IndexError:
            resultexampletable='' 
        isonline = isonline_name(resp.text)
        run_system = run_system_name(resp.text)
        apply_type = re.findall(r"id=\"apply_type\".*?checked>(.*?)<", resp.text, re.U)[-1]
        implement_codenum=implement_code(resp.text)
        service_code = get_service_code(resp.text)
        four_standard=four_standard_name(resp.text)
        onlineHandleDepth=onlineHandleDepth_name(resp.text)
        enterprise_zt=enterprise_zt_name(resp.text)
        person_zt=person_zt_name(resp.text)
        for_user=for_user_name(resp.text)
        asa_unid_text = asa_unid(resp.text)
        isEntryCenter=isEntryCenter_name(resp.text)
        service_type=service_type_name(resp.text)
        info_type = re.findall('(?s)id="info_type".*value=\'(.*?)\' selected', resp.text, re.U)[0]
        content={
            '办理单位':dept_name,
            "星级评定":star_level,
            "所属实施清单":catalog_name,
            "业务办理项名称":service_name,
            "事项类型":service_type,
            '实施编码':implement_codenum,
            '事项编码':service_code,
            '承诺时限':promise_days,
            '法定时间':legal_days,
            '面向用户对象':for_user,
            '法人主题分类':enterprise_zt,
            "个人主题分类":person_zt,
            "法律依据" :legal_standard,
            "申请条件":apply_condition_desc.replace('\t', '').replace('\n', '').replace('\r', '').replace(' ',''),
            "办理流程":flow_desc.replace('\t', '').replace('\n', '').replace('\r', '').replace(' ',''),
            "中介服务事项":mediationservices,
            "是否网办":isonline,
            "办件类型":info_type,
            "运行系统":run_system,
            "四办标志":four_standard,
            "网办深度":onlineHandleDepth,
            '网办地址':web_apply_url,
            '系统名称':webapplyname,
            "入驻大厅方式":apply_type,
            '申请方式':asa_unid_text[0],
            '到窗口最多次数':asa_unid_text[1],
            '承诺到窗口最多次数说明':asa_unid_text[2],
            "是否进驻政务实体大厅":isEntryCenter,
            '办理结果名称':result_name,
            "结果获取说明":result_desc,
            "结果样本":resultexampletable.replace('\t', '').replace('\n', '').replace('\r', '').replace(' ',''),
            "流程图":powerflowimg
        }
        print(content)
        contents.append(content)
    else:
        pass


def write_excle(content,savefile):
    df=pd.DataFrame.from_dict(content)
    df.set_index(df.columns[0],inplace=True)
    df.to_excel(savefile)



if __name__ == "__main__":
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
    headers = {
        'user-agent': 'Mozilla/5.0 (Windows NT 6.1;'
                      ' Win64; x64) AppleWebKit/537.36 '
                      '(KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36'}
    contents = []
    #设置文件保存的地址
    filename=".\date\修武县政务服务发布库链接.xlsx"
    savefile=".\date\修武县政务事项详情.xlsx"
    lp=LoginPage(driver)
    lp.to_login().get_datebasepage()
    he=HandleExcle(filename)
    he.get_url()
    write_excle(contents,savefile)
    print("{}保存完毕".format(savefile))
    driver.quit()
