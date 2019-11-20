# coding:utf-8
import os
import xlrd
import xlwt

class GetBaseDataObject(object):
    """获取存储接口协议及属性信息"""
    def __init__(self, file_path, testcasefile_path):
        self.base_file_path = file_path
        self.casefile_path = testcasefile_path
    #将两个列表合成一个，样式:XX字段为XX
    def merge_value_exchange_list(self, list_one, list_two):
        merge_list = []
        for i in range(0, len(list_one)):
            if i < len(list_one):
                merge_value = list_one[i] + "字段为：" + list_two[i] + "\n"
                merge_list.append(merge_value)
            else:
                merge_value = list_one[i] + "字段为：" + list_two[i]
                merge_list.append(merge_value)
        return merge_list
    #将两个列表合成一个字典
    def merge_value_exchange_dict(self, list_one, list_two):
        merge_dict = {}
        for i in range(0, len(list_one)):
            merge_dict[list_one[i]] = list_two[i]
        return merge_dict
    #按照字段名称，筛选
    def exclude_value(self, ziduan_name, ziduan_list):
        exclude_list = []
        for i in ziduan_list:
            if ziduan_name in i:
                continue
            else:
                exclude_list.append(i)
        return exclude_list
    #打开保存用例的excel表,之所以要分开写是因为要写多个sheet页，如果每次都重新打开一个表就只能保存最后一组测试用例
    def open_excel(self):
        self.workbook = xlwt.Workbook(encoding='utf-8')
        style = xlwt.XFStyle()  # 初始化样式
        font = xlwt.Font()  # 为样式创建字体
        font.name = 'Times New Roman'
    #将测试用例写入到excel表中
    def write_excel(self, sheet_name, testcase_list):
        worksheet = self.workbook.add_sheet(sheet_name)
        worksheet.write(0, 0, label='测试用例目的')
        worksheet.write(0, 1, label='测试用例步骤')
        worksheet.write(0, 2, label='测试用例预期结果')
        worksheet.write(0, 3, label='测试用例评估准则')
        for i in range(1, (len(testcase_list)+1)):
            for j in range(0, 4):
                worksheet.write(i, j, label=testcase_list[i - 1][j])
    #保存并关闭excel
    def save_excel(self):
        self.workbook.save(self.casefile_path)

    def creat_testcase(self):
        data = xlrd.open_workbook(self.base_file_path)
        names = data.sheet_names()#获取表中所有sheet页名字
        self.open_excel()
        for i in names:
            table = data.sheet_by_name(i)
            nrows = table.nrows#获取当前sheet页中的总行数
            testcase_suit = []  # 测试用例集
            self.normal_value_list = self.merge_value_exchange_list(table.col_values(0, start_rowx=4, end_rowx=(nrows + 1)),table.col_values(4, start_rowx=4,end_rowx=(nrows + 1)))  # 用于存放所有字段为正常值的列表
            interface_type = table.cell_value(1, 0)#接口方向
            interface_sim_name = table.cell_value(1, 1)#接口模拟器名称
            software_by_test_name = table.cell_value(1, 2)#被测软件名称
            interface_name = table.cell_value(1, 3)#接口名称
            flag_value = table.cell_value(1, 4)#分隔符
            decimal_format = table.cell_value(1, 5)#进制
            if interface_type == "输入接口":
                testcase = []  # 存储一条测试用例，然后转存至用例集中
                #创建均为正常值的输入用例
                testmudi = "正常值-#@%-%@#-" + "输入" + interface_name + "报文" + "-各字段均为正常值"
                testbuzhou = "使用" + interface_sim_name + "向" + software_by_test_name + "发送" + "“" + interface_name + "”" "具体内容如下" + "" \
                             ":\n" + ("".join(self.merge_value_exchange_list(table.col_values(0, start_rowx=4, end_rowx=nrows + 1), table.col_values(4, start_rowx=4, end_rowx=nrows + 1))))
                testyuqi = software_by_test_name + "可以接收处理" + interface_name + "报文的各字段为正常值，并且正常接收该" + interface_name + "报文"
                testzhunze = software_by_test_name + "可以正确识别" + interface_name + "接口数据，并做出正确处理"
                #将正常值用例各属性添加到一条测试用例列表中
                testcase.append(testmudi)
                testcase.append(testbuzhou)
                testcase.append(testyuqi)
                testcase.append(testzhunze)
                # 将正常值用例列表添加到测试用例集中
                testcase_suit.append(testcase)
                #开始对每个字段的各类值进行测试用例绘制
                for k in range(4, nrows):
                    ziduan_value_dict = self.merge_value_exchange_dict(table.row_values(3, start_colx=5, end_colx=12), table.row_values(k, start_colx=5, end_colx=12))
                    #开始写用例
                    for l in ziduan_value_dict:
                        if l in ["有效范围值", "枚举值", "业务边界值", "据字段长度的边界值"]:#圈定正常和异常范围
                            if ziduan_value_dict[l] == "":#如果字段没有值就继续循环，而不是用break跳出循环喔
                                continue
                            else:
                                ziduan_value_suit = ziduan_value_dict[l].split("\n")
                                if len(ziduan_value_suit) > 1:#正常值的用例，存在多种情况那种
                                    for m in ziduan_value_suit:
                                        testcase = []  # 存储一条测试用例，然后转存至用例集中
                                        testmudi = "特殊值-#@%-%@#-" + "输入" + interface_name + "报文" + "-" + table.cell_value(k, 0) + "字段为" + l
                                        testbuzhou = "使用" + interface_sim_name + "向" + software_by_test_name + "发送" + "“" + interface_name + "”，具体内容如下:\n" + ("".join(self.exclude_value(table.cell_value(k, 0), self.normal_value_list))) + table.cell_value(k, 0) + "字段为：" + m
                                        testyuqi = software_by_test_name + "可以接收处理" + interface_name + "报文的" + table.cell_value(k, 0) + "字段为" + l + "并且正常接收该" + interface_name + "报文"
                                        testzhunze = software_by_test_name + "可以正确识别" + interface_name + "接口数据，并做出正确处理"
                                        # 将用例各属性添加到一条测试用例列表中
                                        testcase.append(testmudi)
                                        testcase.append(testbuzhou)
                                        testcase.append(testyuqi)
                                        testcase.append(testzhunze)
                                        # 将正常值用例列表添加到测试用例集中
                                        testcase_suit.append(testcase)
                                        #print("测试目的：%s\n测试步骤：%s\n测试预期：%s\n测试准则：%s\n******************" % (testmudi, testbuzhou, testyuqi, testzhunze))
                                else:#正常值的用例，不存在多种情况那种
                                    testcase = []  # 存储一条测试用例，然后转存至用例集中
                                    testmudi = "特殊值-#@%-%@#-" + "输入" + interface_name + "报文" + "-" + table.cell_value(k, 0) + "字段为" + l
                                    testbuzhou = "使用" + interface_sim_name + "向" + software_by_test_name + "发送" + "“" + interface_name + "”，具体内容如下:\n" + ("".join(self.exclude_value(table.cell_value(k, 0),self.normal_value_list))) + table.cell_value(k, 0) + "字段为：" + "".join(ziduan_value_suit)
                                    testyuqi = software_by_test_name + "可以接收处理" + interface_name + "报文的" + table.cell_value(k, 0) + "字段为" + l + "并且正常接收该" + interface_name + "报文"
                                    testzhunze = software_by_test_name + "可以正确识别" + interface_name + "接口数据，并做出正确处理"
                                    # 将用例各属性添加到一条测试用例列表中
                                    testcase.append(testmudi)
                                    testcase.append(testbuzhou)
                                    testcase.append(testyuqi)
                                    testcase.append(testzhunze)
                                    # 将正常值用例列表添加到测试用例集中
                                    testcase_suit.append(testcase)
                        else:#异常值的用例
                            if ziduan_value_dict[l] == "":#如果字段没有值就继续循环，而不是用break跳出循环喔
                                continue
                            else:
                                ziduan_value_suit = ziduan_value_dict[l].split("\n")
                                if len(ziduan_value_suit) > 1:#异常值的用例，存在多种情况那种
                                    for n in ziduan_value_suit:
                                        testcase = []  # 存储一条测试用例，然后转存至用例集中
                                        testmudi = "异常值-#@%-%@#-" + "输入" + interface_name + "报文" + "-" + table.cell_value(k, 0) + "字段为" + l
                                        testbuzhou = "使用" + interface_sim_name + "向" + software_by_test_name + "发送" + "“" + interface_name + "”，具体内容如下:\n" + ("".join(self.exclude_value(table.cell_value(k, 0), self.normal_value_list))) + table.cell_value(k, 0) + "字段为：" + n
                                        testyuqi = software_by_test_name + "可以接收处理" + interface_name + "报文的" + table.cell_value(k, 0) + "字段为" + l + "并且抛出该报文"
                                        testzhunze = software_by_test_name + "可以正确识别" + interface_name + "接口数据，并做出正确处理"
                                        # 将用例各属性添加到一条测试用例列表中
                                        testcase.append(testmudi)
                                        testcase.append(testbuzhou)
                                        testcase.append(testyuqi)
                                        testcase.append(testzhunze)
                                        # 将正常值用例列表添加到测试用例集中
                                        testcase_suit.append(testcase)
                                        #print("测试目的：%s\n测试步骤：%s\n测试预期：%s\n测试准则：%s\n******************" % (testmudi, testbuzhou, testyuqi, testzhunze))
                                else:#异常值的用例，不存在多种情况那种
                                    testcase = []  # 存储一条测试用例，然后转存至用例集中
                                    testmudi = "异常值-#@%-%@#-" + "输入" + interface_name + "报文" + "-" + table.cell_value(k, 0) + "字段为" + l
                                    testbuzhou = "使用" + interface_sim_name + "向" + software_by_test_name + "发送" + "“" + interface_name + "”，具体内容如下:\n" + ("".join(self.exclude_value(table.cell_value(k, 0),self.normal_value_list))) + table.cell_value(k, 0) + "字段为：" + "".join(ziduan_value_suit)
                                    testyuqi = software_by_test_name + "可以接收处理" + interface_name + "报文的" + table.cell_value(k, 0) + "字段为" + l + "并且抛出该报文"
                                    testzhunze = software_by_test_name + "可以正确识别" + interface_name + "接口数据，并做出正确处理"
                                    # 将用例各属性添加到一条测试用例列表中
                                    testcase.append(testmudi)
                                    testcase.append(testbuzhou)
                                    testcase.append(testyuqi)
                                    testcase.append(testzhunze)
                                    # 将正常值用例列表添加到测试用例集中
                                    testcase_suit.append(testcase)
                self.write_excel(i + interface_type, testcase_suit)
            elif interface_type == "输出接口":
                testcase = []  # 存储一条测试用例，然后转存至用例集中
                # 创建均为正常值的输入用例
                testmudi = "正常操作-#@%-%@#-" + "输出" + interface_name + "报文"
                testbuzhou = "通过功能驱动，使" + software_by_test_name + "向" + interface_sim_name + "输出" + "“" + interface_name + "”报文，检查报文样式、内容是否与接口文档所述一致"
                testyuqi = software_by_test_name + "输出的" + interface_name + "报文，其样式符合文档要求，内容与功能操作匹配"
                testzhunze = software_by_test_name + "具备输出符合接口文档要求的" + interface_name + "报文"
                # 将正常值用例各属性添加到一条测试用例列表中
                testcase.append(testmudi)
                testcase.append(testbuzhou)
                testcase.append(testyuqi)
                testcase.append(testzhunze)
                # 将正常值用例列表添加到测试用例集中
                testcase_suit.append(testcase)
                self.write_excel(i + interface_type, testcase_suit)
            else:
                print("用例数据表的接口方向有错误！")
        self.save_excel()


if __name__ == "__main__":
    InterfaceCaseCreat_file_realpath = os.path.realpath(__file__)
    InterfaceCaseCreat_file_ospath = os.path.dirname(InterfaceCaseCreat_file_realpath)
    workcode_ospath = os.path.dirname(InterfaceCaseCreat_file_ospath)
    InterfaceTestCase_Base_File_realpath = os.path.join(workcode_ospath, "Database", "InterfaceTestCase_Base_File.xlsx")
    InterfaceTestCase_File_realpath = os.path.join(workcode_ospath, "Database", "InterfaceTestCase_File.xls")
    getbasedata = GetBaseDataObject(InterfaceTestCase_Base_File_realpath, InterfaceTestCase_File_realpath)
    getbasedata.creat_testcase()