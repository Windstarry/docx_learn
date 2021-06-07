import pandas as pd
import os
from utils.config import FILENAME,FILESAVEPATH

class GovernmentAffair(object):

    def __init__(self,num):
        self.filename = FILENAME
        self.num = num
        self.df = pd.read_excel(self.filename)
        self.implement_codenum = self.get_implement_codenum()
        self.service_codenum = self.get_service_codenum()
        self.dept_name = self.get_dept_name()
        self.service_name = self.get_service_name()
        self.for_user = self.get_for_user()
        self.service_type = self.get_service_type()
        self.apply_condition_desc = self.get_apply_condition_desc()
        self.legal_foundation = self.get_legal_foundation()
        self.flow_desc = self.get_flow_desc()
        self.flow_type = self.get_flow_type()
        self.result_name = self.get_result_name()
        self.resultexampletable = self.get_resultexampletable()
        self.accept_window = self.get_accept_window()
        self.legal_days = self.get_legal_days()
        self.promise_days = self.get_promise_days()
        self.consult_tel = self.get_consult_tel()
        self.complaint_tel = self.get_complaint_tel()
        self.mediation_service = self.get_mediation_service()
        self.materia_lists = self.get_materia_lists()
        self.service_unid = self.get_service_unid()
        self.dept_unid = self.get_dept_unid()
  

    def get_implement_codenum(self):
        implement_codenum=self.df['实施编码'].at[self.num]
        return implement_codenum


    def get_service_codenum(self):
        service_codenum=self.df['事项编码'].at[self.num]
        return service_codenum


    def get_dept_name(self):
        dept_name = self.df['办理单位'].at[self.num]
        return dept_name


    def get_service_name(self):
        service_name = self.df['业务办理项名称'].at[self.num]
        return service_name


    def get_for_user(self):
        for_user = self.df['面向用户对象'].at[self.num]
        return for_user


    def get_service_type(self):
        service_type = self.df['事项类型'].at[self.num]
        return service_type


    def get_apply_condition_desc(self):
        apply_condition_desc = self.df['申请条件'].at[self.num]
        return apply_condition_desc


    def get_legal_foundation(self):
        legal_foundation= self.df['法律依据'].at[self.num]
        return legal_foundation


    def get_flow_desc(self):
        flow_desc = self.df['办理流程'].at[self.num]
        return flow_desc


    def get_flow_type(self):
        flow_type = self.df['办理流程类型'].at[self.num]
        return flow_type


    def get_result_name(self):
        if pd.isnull(self.df['办理结果名称'].at[self.num]):
            result_name = '无办理结果'
        else:
            result_name = self.df['办理结果名称'].at[self.num]
        return result_name


    def get_accept_window(self):
        accept_window = self.df['受理窗口'].at[self.num]
        return accept_window


    def get_resultexampletable(self):
        if pd.isnull(self.df['结果样本'].at[self.num]):
            resultexampletable = '无结果样本'
        else:
            resultexampletable = self.df['结果样本'].at[self.num]
        return resultexampletable

    
    def get_legal_days(self):
        legal_days = self.df['法定时间'].at[self.num]
        return legal_days

    
    def get_promise_days(self):
        promise_days = self.df['承诺时限'].at[self.num]
        return promise_days


    def get_consult_tel(self):
        consult_tel = self.df['咨询电话'].at[self.num]
        return consult_tel


    def get_complaint_tel(self):
        complaint_tel = self.df['投诉电话'].at[self.num]
        return complaint_tel


    def get_mediation_service(self):
        mediation_service = self.df['中介服务事项'].at[self.num]
        return mediation_service

    
    def get_materia_lists(self):
        materia_list = self.df['材料清单'].at[self.num]
        return materia_list

    
    def get_service_unid(self):
        service_unid = self.df['unid'].at[self.num]
        return service_unid

    
    def get_dept_unid(self):
        dept_unid = self.df['dept_unid'].at[self.num]
        return dept_unid


class MaterialList(object):
    
    def __init__(self,material_num,material_service_name,material_service_unid,material_dept_name,material_name,material_file_url):
        self.num = material_num
        self.service_name = material_service_name
        self.service_unid = material_service_unid
        self.dept_name = material_dept_name
        self.name = material_name
        self.file_url = material_file_url
        self.material_file_url = self.get_material_file_url()
        self.service_name_path = self.material_path()
        self.material_save_file = self.material_save_file_name()


    def get_material_file_url(self):
        base_url = 'https://uploadmatter.hnzwfw.gov.cn/fileserver/download.jsp?filePath='
        material_file_url = ''.join([base_url,f'{self.file_url}'])
        return  material_file_url


    def material_save_path(self):
        result_path = FILESAVEPATH
        service_name = self.service_name.replace('/','、').replace('<','').replace('>','')
        dept_path = os.path.join(result_path, f'{self.dept_name}')
        if not os.path.exists(dept_path):
            os.mkdir(dept_path)
        service_name_path = os.path.join(dept_path, f'{service_name}')
        if not os.path.exists(service_name_path):
            os.mkdir(service_name_path)

    
    def material_path(self):
        result_path = FILESAVEPATH
        service_name = self.service_name.replace('/','、').replace('<','').replace('>','')
        dept_path = os.path.join(result_path, f'{self.dept_name}')
        service_name_path = os.path.join(dept_path, f'{service_name}')
        return service_name_path


    def material_save_file_name(self):
        file_format = self.file_url.split('.')[1]
        save_name = '{}-申请材料{}-{}'.format(self.service_name,self.num,self.name)
        file_save_name = '.'.join([save_name,file_format])
        material_save_file = os.path.join(f'{self.service_name_path}',file_save_name)
        return material_save_file


class MaterialExampleList(object):
    
    def __init__(self,material_num,material_service_name,material_service_unid,material_dept_name,material_name,material_example_fileurl):
        self.num = material_num
        self.service_name = material_service_name
        self.service_unid = material_service_unid
        self.dept_name = material_dept_name
        self.name = material_name
        self.exampleF_file_url = material_example_fileurl
        self.material_example_file_url = self.get_material_example_file_url()
        self.service_name_path = self.material_path()
        self.material_example_save_file = self.material_example_save_file_name()

        
    def get_material_example_file_url(self):
        base_url = 'https://uploadmatter.hnzwfw.gov.cn/fileserver/download.jsp?filePath='
        material_example_file_url = ''.join([base_url,self.exampleF_file_url])
        return material_example_file_url


    def material_save_path(self):
        result_path = FILESAVEPATH
        service_name = self.service_name.replace('/','、').replace('<','').replace('>','')
        dept_path = os.path.join(result_path, f'{self.dept_name}')
        if not os.path.exists(dept_path):
            os.mkdir(dept_path)
        service_name_path = os.path.join(dept_path, f'{service_name}')
        if not os.path.exists(service_name_path):
            os.mkdir(service_name_path)

    
    def material_path(self):
        result_path = FILESAVEPATH
        service_name = self.service_name.replace('/','、').replace('<','').replace('>','')
        dept_path = os.path.join(result_path, f'{self.dept_name}')
        service_name_path = os.path.join(dept_path, f'{service_name}')
        return service_name_path

   
    def material_example_save_file_name(self):
        file_format = self.exampleF_file_url.split('.')[1]
        save_name = '{}-申请材料{}-{}-示例文本'.format(self.service_name,self.num,self.name)
        example_file_save_name = '.'.join([save_name,file_format])
        material_example_save_file = os.path.join(f'{self.service_name_path}',example_file_save_name)
        return material_example_save_file



    
