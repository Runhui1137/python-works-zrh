import docx
import os
from win32com import client
import csv
import re

# 配置信息
docx_path = "D:/2020春课堂/DOTNET/实验1/182实验报告"       # 目标文件夹的路径
csv_path = './182.csv'                                  # csv文件路径, 用于录入学生学号等信息
reg_patt = u'2018\d\d\d\d\d\d'                          # 正则表达式, 用于在docx中查找学号
# 要格式化实验报告文件名, 直接修改上免变量的值即可
#########################################################################
word = client.Dispatch("Word.Application")    # 用于doc格式转换为docx的工具
student_map = dict()                          # 学生字典, 学号为key, 用于存储学生的信息, 这些信息将用于生成新的文件名

def get_files():
    return os.listdir(docx_path)

# modify_single_file: 对单个文件进行改名
## filename: 文件名(不含路径)
def modify_single_file(filename):
    # 判断并记录该文件的类型
    is_docx = True
    if (filename.split(".")[-1] == 'doc'):
        is_docx = False
    elif (filename.split(".")[-1] == 'docx'):
        is_docx = True
    else:
        return
    # 得到该文件的路径
    if docx_path[-1] != '\\' :
        filepath = docx_path + '\\'
    filepath += filename
    docx_filepath = filepath     # 变量docx_filepath是为了解决doc文件转换而存在, 如果该文件为docx格式, 该变量与filepath值相同
    extension = 'docx'          # 变量extension即为文件之前的扩展名(doc或docx)
    # 如果不为docx文件, 则生成临时docx文件
    if not is_docx :
        docx_filepath = save_single_doc_as_docx(filepath, filename)
        extension = 'doc'
    # 打开并遍历docx, 使用正则表达式进行匹配学号
    re_ans = None
    dx = docx.Document(docx=docx_filepath)
    for line in dx.paragraphs :
        re_ans = re.search(string=line.text, pattern=reg_patt)
        if re_ans != None:
            break
    # 若学号读取成功, 则通过student_map得到该学生的信息, 生成目标文件名, 对源文件完成重命名工作
    new_name = ''
    if re_ans != None:
        sno = re_ans.group()
        info = student_map[sno]
        ###### 在这里可以修改重命名文件的格式
        new_name = "{0} {1}.{2}".format(info[1], info[0], extension)
        new_path = docx_path+ "\\" + new_name
        os.rename(filepath, new_path)
    # 删除临时生成的docx文件
    if not is_docx:
        os.remove(docx_filepath)
    print("{0}    --->    {1}      success !".format(filename, new_name))

# 将doc文件另存为docx文件至temp目录中
def save_single_doc_as_docx(filepath, filename):
    docx_filepath = docx_path + '\\temp\\' + filename.split('.')[0] + '.docx'
    doc = word.Documents.Open(filepath)
    doc.SaveAs(docx_filepath, 12)
    doc.Close()
    return docx_filepath

# 从CSV文件中读取并录入学生信息到student_map中
def get_info():
    with open(csv_path) as file:
        for item in csv.reader(file) :
            t = student_map[item[0]] = item[0:2]    # 根据我这个csv文件, item[0]即为学号

def main():
    global docx_path
    docx_path = os.path.abspath(docx_path)
    docx_path = "\\".join(docx_path.split('/'))
    if not os.path.exists(docx_path + "\\temp"):
        os.mkdir(docx_path + "\\temp")
    # 获取目标文件夹下的所有文件
    print(docx_path)
    file_list = get_files()
    # 获取学生信息(这些信息后来将用于生成文件名)
    get_info()
    # 对每个文件重命名
    for file in file_list:
        modify_single_file(file)
    print("All Tasks Success!")
    # 析构win32com
    # 在没使用完win32com模块前千万别调用Quit方法, 否则会出现错误
    word.Quit()

main()
