"""
该脚本用作单次根据台账统计所需要的数据
"""

import openpyxl
import numpy as np
import pandas as pd
import xlrd
import datetime


class Student:
    # status: int 0大陆学生 1香港学生 2台湾学生 3外籍学生
    # name: str
    # id_number: str
    # student_number: str
    # age: int，仅大陆学生有
    def __init__(self, name, id_number, student_number, status):
        self.name = name
        self.id_number = id_number
        self.student_number = student_number
        self.status = status
        if status == 0:
            birthday = id_number[6:14]
            stu_year = int(birthday[:4])
            stu_month = int(birthday[4:6])
            stu_day = int(birthday[6:])
            now_year = int(datetime.datetime.now().year)
            now_month = int(datetime.datetime.now().month)
            now_day = int(datetime.datetime.now().day)
            if stu_month > now_month or (stu_month == now_month and stu_day >= now_day):
                age = now_year - stu_year
            else:
                age = now_year - stu_year - 1
            self.age = age
        else:
            self.age = -1

    def show_info(self):
        print(self.name + '   身份:' + str(self.status) + '   ' + self.id_number + '   ' + self.student_number)


name_list = pd.read_excel('台账.xls', dtype=str)
test_data = pd.read_excel('核酸数据.xlsx', dtype=str)
info_list = pd.read_excel('0717.xlsx', dtype=str)

'''洗去制表符'''
for i in range(test_data.shape[0]):
    if test_data.iat[i, 2][-1] == '\t':
        test_data.iat[i, 2] = test_data.iat[i, 2][:-1]

for i in range(name_list.shape[0]):
    if name_list.iat[i, 1][-1] == '\t':
        name_list.iat[i, 1] = name_list.iat[i, 1][:-1]

'''三个表字符串姓名全小写'''
for i in range(name_list.shape[0]):
    name_list.iat[i, 2] = str(name_list.iat[i, 2]).lower()

for i in range(info_list.shape[0]):
    info_list.iat[i, 2] = str(info_list.iat[i, 2]).lower()

for i in range(test_data.shape[0]):
    test_data.iat[i, 3] = str(test_data.iat[i, 3]).lower()

'''将info_list以姓名为主去除重复数据（以最后一次为准），并将台账中不存在的数据去除'''
info_list = info_list.drop_duplicates(subset=['Q1_姓名'], keep='last', inplace=False)

s = name_list['学号'].values
n = name_list['姓名'].values
to_be_delete = []
for i in range(info_list.shape[0]-1, -1, -1):
    if not (info_list.iat[i, 3] in s or info_list.iat[i, 2] in n):
        to_be_delete.append(i)
info_list.index = range(info_list.shape[0])
info_list = info_list.drop(index=to_be_delete)
info_list.index = range(info_list.shape[0])

# info_list = info_list.reindex(range(len(info_list)))

''' 
登记个人信息，前提是info_list已完成去重
'''
stu_info = []
for i in range(info_list.shape[0]):
    # 根据具体位置记录原始数据
    name = str(info_list.iat[i, 2])
    id_number = str(info_list.iat[i, 4])
    stu_number = str(info_list.iat[i, 3])
    status = str(info_list.iat[i, 14])
    # status: int 0大陆学生 1香港学生 2台湾学生 3外籍学生
    if status == '香港':
        status = 1
    else:
        if status == '台湾':
            status = 2
        else:
            if status == '外籍':
                status = 3
            else:
                if status == '大陆':
                    status = 0
                else:
                    status = -1

    temp = Student(name, id_number, stu_number, status)
    stu_info.append(temp)

'''
重名逻辑未完成
'''

'''以姓名为主去除重复数据（以最后一次为准），将核酸数据中在台账中不存在的学生去除，将疫苗信息重新设置索引'''
test_data = test_data.drop_duplicates(subset=['Q2_姓名'], keep='last', inplace=False)

s = name_list['学号'].values
n = name_list['姓名'].values
to_be_delete = []
for i in range(test_data.shape[0]-1, -1, -1):
    if not (test_data.iat[i, 2] in s or test_data.iat[i, 3] in n):
        to_be_delete.append(i)
test_data.index = range(test_data.shape[0])
test_data = test_data.drop(to_be_delete)
test_data.index = range(test_data.shape[0])

# test_data = test_data.reindex(range(len(test_data)))

'''按台账校正学号未完成'''

print(name_list)
print(test_data)



to_be_delete = []
for i in range(name_list.shape[0]):
    if name_list.iat[i, 6] != '大陆学生':
        to_be_delete.append(i)
name_list.index = range(name_list.shape[0])
mainland = name_list.drop(to_be_delete)
mainland.index = range(mainland.shape[0])
xuanmin_list = pd.read_excel('选民数据.xlsx', dtype = str)
xuanmin_name = xuanmin_list['*姓名'].values

all_stu = pd.read_excel('全院.xls', dtype = str)
for i in range(mainland.shape[0]):
    if not mainland.iat[i, 2] in xuanmin_name:
        index = -1
        for j in range(all_stu.shape[0]):
            if all_stu.iat[j, 0] == mainland.iat[i, 2]:
                index = j
                break
        if index == -1:
            print('不在选民数据但在台账中的\t'+mainland.iat[i, 2] + '\t在全院数据中未找到')
        else:
            status = '默认'
            if all_stu.iat[index, 2][0:1]=='0':
                status = '博士研究生'
            else:
                if all_stu.iat[index, 2][0:1]=='5':
                    status = '普通本科生'
            try:
                print(all_stu.iat[index, 2]+'\t'+all_stu.iat[index, 0]+'\t'+all_stu.iat[index, 12][:1]+'\t'+'数学科学学院'+'\t'+all_stu.iat[index,3]+'\t'+status+'\t'+all_stu.iat[index, 6] + '\t' + all_stu.iat[index, 11])
            except:
                print('error!' + all_stu.iat[index, 0])
                print(all_stu.iat[index, 2])
                print(all_stu.iat[index, 12][:1])
                print(all_stu.iat[index, 3])
                print(status)
                print(all_stu.iat[index, 6])

'''
统计小脚本
'''
'''
under_18 = ['温晓烽', '刁守淳', '张世博', '石涵', '王湛']
adult_all = 0
adult_2 = 0
adult_1 = 0
adult_no = 0
adult_only1 = 0
teen_all = 0
teen_2 = 0
teen_1 = 0
teen_no = 0
teen_only1 = 0
for i in range(test_data.shape[0]):
    if test_data.iat[i, 7] == '大陆学生' and (test_data.iat[i, 3] in under_18):
        teen_all = teen_all + 1
        if test_data.iat[i, 29] == '尚未接种（如无禁忌症，请务必在入校之前接种疫苗）':
            teen_no = teen_no + 1
        else:
            if test_data.iat[i, 29] == '已经完成接种两针剂疫苗（总共需要接种两针）':
                teen_2 = teen_2 + 1
            else:
                if test_data.iat[i, 29] == '已经完成接种第一针疫苗（总共需要接种两针）':
                    teen_1 = teen_1 + 1
                else:
                    if test_data.iat[i, 29] == '已经完成接种一次性针剂疫苗（总共需要接种一针）':
                        teen_only1 = teen_only1 + 1
                    else:
                        print('疫苗信息error! ' + test_data.iat[i, 3])
    else:
        if test_data.iat[i, 7] == '大陆学生' and (not (test_data.iat[i, 3] in under_18)):
            adult_all = adult_all + 1
            if test_data.iat[i, 29] == '尚未接种（如无禁忌症，请务必在入校之前接种疫苗）':
                adult_no = adult_no + 1
            else:
                if test_data.iat[i, 29] == '已经完成接种两针剂疫苗（总共需要接种两针）':
                    adult_2 = adult_2 + 1
                else:
                    if test_data.iat[i, 29] == '已经完成接种第一针疫苗（总共需要接种两针）':
                        adult_1 = adult_1 + 1
                    else:
                        if test_data.iat[i, 29] == '已经完成接种一次性针剂疫苗（总共需要接种一针）':
                            adult_only1 = adult_only1 + 1
                        else:
                            print('疫苗信息error! ' + test_data.iat[i, 3])
'''
'''
print('满18周岁大陆学生总数：' + str(adult_all))
print('满18周岁大陆学生已接种两针（共两针）人数：' + str(adult_2))
print('满18周岁大陆学生已接种一针（共两针）人数：' + str(adult_1))
print('满18周岁大陆学生已接种一针（共一针）人数：' + str(adult_only1))
print('满18周岁大陆学生未接种疫苗人数：' + str(adult_no))
print()
print()
print('未满18周岁大陆学生总数：' + str(teen_all))
print('未满18周岁大陆学生已接种两针（共两针）人数：' + str(teen_2))
print('未满18周岁大陆学生已接种一针（共两针）人数：' + str(teen_1))
print('未满18周岁大陆学生已接种一针（共一针）人数：' + str(teen_only1))
print('未满18周岁大陆学生未接种疫苗人数：' + str(teen_no))
'''






'''
# 筛选中国大陆学生

mainland_stu_info = []
for i in range(len(stu_info)):
    if stu_info[i].status == 0:
        mainland_stu_info.append(stu_info[i])

# 以name为主键使用问卷数据统计大陆学生成年人及未成年人疫苗接种情况

adult_all = 0
adult_2 = 0
adult_1 = 0
adult_no = 0
adult_already = 0
teen_all = 0
teen_2 = 0
teen_1 = 0
teen_no = 0
teen_already = 0
no_info_num = 0
no_info_index = []
print(test_data.shape[0])
for i in range(test_data.shape[0]):
    index = -1
    for j in range(len(mainland_stu_info)):
        if test_data.iat[i, 3] == mainland_stu_info[j].name:
            index = j
    if index != -1:
        if mainland_stu_info[index].age >= 18:
            adult_all = adult_all + 1
            if str(test_data.iat[i, 29]) == '已经完成接种第一针疫苗（总共需要接种两针）':
                adult_1 = adult_1 + 1
            else:
                if str(test_data.iat[i, 29]) == '已经完成接种两针剂疫苗（总共需要接种两针）':
                    adult_2 = adult_2 + 1
                else:
                    adult_no = adult_no + 1
        if mainland_stu_info[index].age < 18:
            teen_all = teen_all + 1
            if str(test_data.iat[i, 29]) == '已经完成接种第一针疫苗（总共需要接种两针）':
                teen_1 = teen_1 + 1
            else:
                if str(test_data.iat[i, 29]) == '已经完成接种两针剂疫苗（总共需要接种两针）':
                    teen_2 = teen_2 + 1
                else:
                    teen_no = teen_no + 1
    else:
        print('在学生信息中未找到 ' + str(test_data.iat[i, 3]) + '  ' + str(test_data.iat[i, 2]))
        no_info_index.append(i)
        no_info_num = no_info_num + 1

adult_already = adult_all - adult_no
teen_already = teen_all - teen_no

print('学生总人数：' + str(name_list.shape[0]))
print('学生信息总人数：' + str(info_list.shape[0]))
print('疫苗信息总人数：' + str(test_data.shape[0]))
print('无身份证信息总人数：' + str(no_info_num) + '  注：无身份证信息在test_data中的index信息保存在no_info_index中')
'''

'''
检查台账中所有人学号或姓名是否在核酸报告中出现过，没出现过的输出学号及姓名
s = test_data['Q1_学号'].values
n = test_data['Q2_姓名'].values
for i in range(name_list.shape[0]):
    if not (name_list.iat[i,1] in s or name_list.iat[i,2] in n):
        print(str(name_list.iat[i,1])+'   '+str(name_list.iat[i,2]))
'''

'''
反过来检查台账中所有人学号或姓名是否在核酸报告中出现过，没出现过的输出学号及姓名
s = name_list['学号'].values
n = name_list['姓名'].values
for i in range(test_data.shape[0]):
    if not (test_data.iat[i,2] in s or test_data.iat[i,3] in n):
        print(str(test_data.iat[i,2])+'   '+str(test_data.iat[i,3]))
'''

'''
检查长度是否为12，并报出所有不是的情况
for i in range(test_data.shape[0]):
    if len(test_data.iat[i, 2])!=12:
        print(str(i)+'  '+str(test_data.iat[i,2]))
'''

'''
检查是否为数（无用）
t = np.int64(1)
for i in range(test_data.shape[0]):
    if test_data.iat[i,2].dtype != t.dtype:
        print(str(i)+'  '+str(test_data.iat[i,2]))
'''
