import requests
import pandas as pd
import time

class GovernmentAffair(object):

    def __init__(self,num):
        self.filename = './date/修武县政务服务发布库链接.xlsx'
        self.num = num
        self.df = pd.read_excel(self.filename)
        self.max_row = self.df.shape[0]
        self.service_codenum = self.get_service_codenum()
        self.dept_name = self.get_dept_name()
        self.service_name = self.get_service_name()
        self.service_unid = self.get_service_unid()
        self.dept_unid = self.get_dept_unid()


    def get_service_codenum(self):
        service_codenum=self.df['事项编码'].at[self.num]
        return service_codenum


    def get_dept_name(self):
        dept_name = self.df['办理单位'].at[self.num]
        return dept_name


    def get_service_name(self):
        service_name = self.df['业务办理项名称'].at[self.num]
        return service_name

    
    def get_service_unid(self):
        service_unid = self.df['unid'].at[self.num]
        return service_unid

    
    def get_dept_unid(self):
        dept_unid = self.df['dept_unid'].at[self.num]
        return dept_unid


def get_materia_list(service_unid):
    base_url = 'https://www.hnzwfw.gov.cn/hnzwfw/matter/service?serviceUnid={}'
    url = base_url.format(service_unid)
    headers = {
            'User-Agent':'Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.77 Safari/537.36 Edg/91.0.864.37',
            }
    resp = requests.get(url= url,headers=headers)
    resp.encoding='utf-8'
    matter_dict = resp.json()
    materia_list = matter_dict.get('result').get('materialInfo').get('materialList')
    return materia_list


def add_materia_list(contents,governmentaffair):
    dept_name = governmentaffair.dept_name
    service_name = governmentaffair.service_name
    service_codenum = governmentaffair.service_codenum
    unid = governmentaffair.service_unid
    dept_unid = governmentaffair.dept_unid
    materia_list = get_materia_list(governmentaffair.service_unid)
    content={
        '办理单位': dept_name,
        "业务办理项名称": service_name,
        "事项编码": service_codenum,
        "unid": unid,
        "dept_unid": dept_unid,
        "材料清单": materia_list,
    }
    print(content)
    contents.append(content)



def get_result_list(service_unid):
    base_url = 'https://www.hnzwfw.gov.cn/hnzwfw/matter/service?serviceUnid={}'
    url = base_url.format(service_unid)
    headers = {
            'User-Agent':'Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.77 Safari/537.36 Edg/91.0.864.37',
            }
    resp = requests.get(url= url,headers=headers)
    resp.encoding='utf-8'
    matter_dict = resp.json()
    result_name = matter_dict.get('result').get('baseInfo').get('resultName')
    result_example = matter_dict.get('result').get('baseInfo').get('resultExample')
    return result_name,result_example


def add_result_list(contents,governmentaffair):
    dept_name = governmentaffair.dept_name
    service_name = governmentaffair.service_name
    service_codenum = governmentaffair.service_codenum
    unid = governmentaffair.service_unid
    dept_unid = governmentaffair.dept_unid
    result_name = get_result_list(governmentaffair.service_unid)[0]
    result_example = get_result_list(governmentaffair.service_unid)[1]
    content={
        '办理单位':dept_name,
        "业务办理项名称":service_name,
        "事项编码":service_codenum,
        "unid":unid,
        "dept_unid":dept_unid,
        "结果名称":result_name,
        "结果样本":result_example,
    }
    print(content)
    contents.append(content)


def write_excle(content,savefile):
    df=pd.DataFrame.from_dict(content)
    df.set_index(df.columns[0],inplace=True)
    df.to_excel(savefile)


if __name__ == '__main__':
    materia_save_file = './date/修武县政务服务发布库材料清单.xlsx'
    result_save_file = './date/修武县政务服务发布库办理结果.xlsx'
    materia_contents = []
    result_contents = []
    for i in range(19,30):
        governmentaffair = GovernmentAffair(i)
    #     add_materia_list(materia_contents,governmentaffair)
    # write_excle(materia_contents,materia_save_file)   
        add_result_list(result_contents,governmentaffair)
    write_excle(result_contents,result_save_file)

    