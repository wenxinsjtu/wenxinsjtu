# 本脚本用以按姓名拼音排序并批量化重命名文件
# 初始格式例：“姓名 *****.pdf”
# 最终格式例：“2 姓名.pdf”

import os
# 用以系统相关操作
from pypinyin import pinyin, Style
# 用以按拼音排序汉字

dir_name = 'C:\\Users\\XinWen\Desktop\\2021-2022春数学分析（荣誉）II（64人）\\答卷'
# 设定要排序的文件所在的文件夹

print('所需排序文件夹位置：' + dir_name)
list_path = os.listdir(dir_name)
print('文件夹中所含文件：')
print(list_path)
# 用以检查文件夹内文件，便于去除可能不需排序的文件

list_path.sort(key=lambda keys:[pinyin(i, style=Style.TONE3) for i in keys])
print(list_path)

dex = 0
for index in list_path:
    dex = dex + 1
    name = index.split(' ')[0]
    # 取需要的姓名
    kid = index.split('.')[-1]
    # 取后缀名
    path = dir_name + '\\' + index
    new_path = dir_name + '\\' + str(dex) + ' ' + name + '.' + kid
    os.rename(path, new_path)
