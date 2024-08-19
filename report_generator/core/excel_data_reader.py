import pandas as pd
from datetime import datetime

class ExcelDataReader:
    def __init__(self):
        self.vulnerabilities = {}
        self.Icp_infos = {}

    '''
    从Excel文件中读取漏洞信息
    '''

    # 从Excel文件中读取漏洞信息，并将其存储到字典中
    def read_vulnerabilities_from_excel(self, file_path):
        data = pd.read_excel(file_path)
        
        vulnerability_names = []  # 存储漏洞名称的列表
        self.vulnerabilities = {}  # 存储漏洞描述和加固建议的字典
        
        for _, row in data.iterrows():
            vulnerability_name = row['漏洞名称']
            vulnerability_description = row['漏洞描述']
            vulnerability_solution = row['加固建议']
            
            vulnerability_names.append(vulnerability_name)  # 将漏洞名称添加到列表中
            
            self.vulnerabilities[vulnerability_name.lower()] = {
                '漏洞描述': vulnerability_description,
                '加固建议': vulnerability_solution
            }
        
        return vulnerability_names, self.vulnerabilities
    
    # 根据漏洞名称从字典中获取漏洞描述和加固建议
    def get_vulnerability_info(self, name):
        name = name.lower()
        if name in self.vulnerabilities:
            description = self.vulnerabilities[name]['漏洞描述']
            solution = self.vulnerabilities[name]['加固建议']
            return description, solution
        else:
            return None, None

    '''
    从Excel文件中读取ICP信息
    '''
	# 从Excel文件中读取ICP信息，并将其存储到字典中
    def read_Icp_from_excel(self, file_path):

        data = pd.read_excel(file_path)
        self.Icp_infos = {}  # 存储单位名称和备案号的字典
        
        for _, row in data.iterrows():
            variables = {
                'domain': row['domain'],
                'unitName': row['unitName'],
                'natureName': row['natureName'],
                'mainLicence': row['mainLicence'],
                'serviceLicence': row['serviceLicence'],
                'updateRecordTime': row['updateRecordTime'],
            }
            # 检查并打印出哪些变量是 NaN, 也就是列表内存在空值, 如果为NaN将其替换为空字符串
            for key, value in variables.items():
                if pd.isna(value):
                    variables[key] = ""

            # 如果 updateRecordTime 不是 NaN，将其转换为日期字符串
            if variables['updateRecordTime'] != "":
                # 判断变量类型并进行相应的转换
                if isinstance(variables['updateRecordTime'], datetime):
                    # 将 datetime 对象转换为字符串
                    variables['updateRecordTime'] = variables['updateRecordTime'].strftime("%Y-%m-%d")
                elif isinstance(variables['updateRecordTime'], str):
                    # 提取日期部分并更新字典
                    variables['updateRecordTime'] = variables['updateRecordTime'].split(" ")[0]
            # 将信息存储到字典中，使用域名作为键
            self.Icp_infos[variables['domain'].lower()] = variables
        return self.Icp_infos
    
    # 根据根域名从字典中获取单位名称和备案号
    def get_Icp_info(self, domain):
        domain_to_search = domain.lower()
        if domain_to_search in self.Icp_infos:
            unitName = self.Icp_infos[domain_to_search]['unitName']
            serviceLicence = self.Icp_infos[domain_to_search]['serviceLicence']
            return unitName, serviceLicence
        else:
            return None, None