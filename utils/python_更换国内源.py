import os
path = os.getenv('APPDATA')   #获取环境变量
data = """[global]
timeout = 6000
index-url = http://pypi.douban.com/simple/ 
trusted-host = pypi.douban.com
"""
#  写入的数据
folder_path = os.path.join(path, "pip")
file_path = os.path.join(path, "pip", "pip.ini")
folder = os.path.exists(folder_path)  #判断文件夹是否存在
if not folder:
    os.mkdir(folder_path)  #创建文件夹
f = open(file_path, 'w')   #写文件
f.write(data)
f.close()