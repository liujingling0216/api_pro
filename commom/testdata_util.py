import os
from commom.excel_util import excelUtil

file_path = os.path.join(os.path.dirname(__file__),'../data/test_case.xlsx')
sheet_name = 'Sheet1'

class testdataUtil():
    def __init__(self,file_path,sheet_name):
        self.file_path = file_path
        self.sheet_name =sheet_name
        self.test_data = excelUtil(file_path, sheet_name).get_value_by_dict()

    def __get_value_to_case_dict_first(self):
        """
        将excelUtil中取出的数据转换成满足框架中的格式第一步
        {'case01': [{'测试用例编号': 'case01', '测试步骤': 'step01', '用例执行': '是'}],
        'case02': [{'测试用例编号': 'case02', '测试步骤': 'step01', '用例执行': '是'},{'测试用例编号': 'case02', '测试步骤': 'step02', '用例执行': '是'}]}
        :return:
        """
        test_case_dic_first = {}
        for row in self.test_data:
            test_case_dic_first.setdefault(row['测试用例编号'],[]).append(row)
        return test_case_dic_first


    def get_value_to_case_dict_second(self):
        """
        将excelUtil中取出的数据转换成满足框架中的格式第二步
        [
        {case_id：'case01'，case_info: [{'测试用例编号': 'case01', '测试步骤': 'step01', '用例执行': '是'}]},
        {case_id：'case02'，case_info:[{'测试用例编号': 'case02', '测试步骤': 'step01', '用例执行': '是'},{'测试用例编号': 'case02', '测试步骤': 'step02', '用例执行': '是'}]}
        ]
        :return:
        """
        test_case_list = []
        for k,v in self.__get_value_to_case_dict_first().items():
            test_case_list_second = {}
            test_case_list_second['case_id']=k
            test_case_list_second['case_info']=v
            test_case_list.append(test_case_list_second)
        return test_case_list



if __name__ == '__main__':
    testdata = testdataUtil(file_path,sheet_name)
    # print(testdata.__get_value_to_case_dict_first())
    print(testdata.get_value_to_case_dict_second())


