# -*- coding: utf-8 -*-
import re, os, glob, shutil
import win32com.client

#Param
target=r'test.txt'
word_ext_list = ['.doc', '.docx'];
text_ext_list = ['.txt', '.log'];
log_file_name = 'trace.log'
temp_file_ext = '.text'
temp_folder_name = 'fm_temp'
pre_word_template = r'[a-zA-Z0-9]+([-_\.][a-zA-Z0-9]+)*'
mid_word_template = r'[a-zA-Z0-9]+([-_][a-zA-Z0-9]+)*'
#pattern =re.compile(r'('+pre_word_template+ r'@' + mid_word_template + r'(\'.[a-zA-Z]{2,3}){1,3})(\b)*')
pattern =re.compile(r'('+pre_word_template+ r'@' + mid_word_template + r'([\.][a-zA-Z]{2,3}){1,3})(\b)*')
#当前目录
cur_path = os.path.abspath('.')
result_file_path = cur_path + "\\result.txt"
temp_folder_path = cur_path + '\\' + temp_folder_name 
log_file_path = cur_path + '\\' + log_file_name
#runtime param
has_word_app = False
has_wps_app = True
word_app = None
wps_app = None
log_file_handler = None
mail_list = [];

#相关方法定义
#阶段跟踪
def step_log(step_key):
    print(step_key)

#判断是否为纯文本文件
def is_test_file(path):
     file_ext_cache = os.path.splitext(path)[1]
     for ext_elem in text_ext_list:
        if ext_elem == file_ext_cache:
            return True
     return False

#判断是否为word文档文件
def is_word_file(path):
    file_ext_cache = os.path.splitext(path)[1]
    for ext_elem in word_ext_list:
        if ext_elem == file_ext_cache:
            return True
    return False

#将检索到的邮箱插入列表中
def add_range_mail_to_list(*mail_info):
    if mail_info <> None and mail_info <> []:
        mail_list.extend(mail_info)#list.insert(len(mail_list), mail_info)

#从纯文本中获取邮箱
def get_mail_form_text(path):
    try:
        file_handler = open(path, 'r')
        get_mail_ret = []
        line = file_handler.readline()
        while line:
            re_ret = re.findall(pattern, line)
            if re_ret <> None and 0 < len(re_ret):
                get_mail_ret = get_mail_ret + trim_list_no_at(*re_ret[0])
            line = file_handler.readline()
        return get_mail_ret
    finally:
        file_handler.close()
    return None

#从word文档中获取邮箱
def get_mail_form_word(path):
    if False == has_word_app:
        return []
    step_log("get_mail_form_word" + path)
    temp_file_name = os.path.basename(path)
    temp_file_path = temp_folder_path + '\\' + temp_file_name + temp_file_ext
    try:
        doc = word_app.Documents.Open(path)
        doc.SaveAs(temp_file_path, 4)
    finally:
        doc.Close()
    if True == os.path.exists(temp_file_path):
        return get_mail_form_text(temp_file_path)
    else:
        return []

def get_mail_form_wps(path):
    if False == has_wps_app:
        return []
    step_log("get_mail_form_wps" + path)
    temp_file_name = os.path.basename(path)
    temp_file_path = temp_folder_path + '\\' + temp_file_name + temp_file_ext
    try:
        doc = wps_app.Documents.Open(path)
        doc.SaveAs(temp_file_path, 4)
    finally:
        doc.Close()
    if True == os.path.exists(temp_file_path):
        return get_mail_form_text(temp_file_path)
    else:
        return []

#剔除没@项
def trim_list_no_at(*table_param):
    ret = []
    for elem in table_param:
        if '@' in str(elem):
            ret.insert(len(ret), elem)
    return ret

#初始化本地环境
def init_local_envir():
    step_log("clean evir")
    if True == os.path.exists(temp_folder_path):
        os.rmdir(temp_folder_path)#os.remove(dirPath)
    os.mkdir(temp_folder_path) 
    if True == os.path.exists(result_file_path):
        os.remove(result_file_path)
    if True == os.path.exists(log_file_path):
        os.remove(log_file_path)
    step_log("cleaning finsh")

#文件扫描跟踪
def scan_trac_log(log_file_handler,path, ret):
    log_file_handler.write("scan:%s,cout:%d,result:%s\n" % (path, len(ret), ret))


step_log("call init")
init_local_envir()
#获取当前目录doc文件列表
#file_list_of_ext1 = glob.glob(cur_path + '\\*.' + file_ext1)
#file_list_of_ext2 = glob.glob(cur_path + '\\*.' + file_ext2)
step_log("initing finish")
step_log("call scanning in " + cur_path)
try:
    all_file_list = glob.glob(cur_path + '\\*')  
    log_file_handler = open(log_file_path, 'w')  
    for file in all_file_list:
        step_log("dealing " + file)
        if is_test_file(file) == True:
            get_mail_ret = get_mail_form_text(file)
            scan_trac_log(log_file_handler, file, get_mail_ret)
            add_range_mail_to_list(*get_mail_ret)
        elif is_word_file(file) == True and (True == has_word_app or True == has_wps_app):
            # if None == word_app and has_word_app:
            #     word_app = win32com.client.Dispatch('Word.Application')
            if None == wps_app and has_wps_app:
                wps_app = win32com.client.Dispatch('kwps.application') 
            #get_mail_ret = get_mail_form_word(file)
            #if None == get_mail_ret or [] == get_mail_form_wps:
            get_mail_ret = get_mail_form_wps(file)
            scan_trac_log(log_file_handler, file, get_mail_ret)
            add_range_mail_to_list(*get_mail_ret)
        else:
            scan_trac_log(log_file_handler, file, [])
            print("invald file:"+file)
finally:
    if None <> log_file_handler:
        step_log("close log_file")
        log_file_handler.close()
    if None <> word_app:
        step_log("close word_app")
        word_app.Quit()
    if None <> wps_app:
        step_log("close wps_app")
        wps_app.Quit()
step_log("scanning finish")

#创建输出文件
f = open(result_file_path, 'w')

#遍历所有文件
# for file in midFileNames2:
#     fileName = os.path.basename(file)
#     context = open(dirPath +'\\' + fileName).read()
#     paragraphs = context.split("\n")
#     print (paragraphs)
#     paragraphNum = len(paragraphs)
#     #遍历段落
#     for i in range(0, paragraphNum):
#         resultList = re.findall(pattern, paragraphs[i])
#         if resultList <> []:
#             for k in range(0, len(resultList)):
#                 #可在此做验证网址
#                 f.write(resultList[k][0] + ";")

for mail_elem in mail_list:
    print(mail_elem)
    f.write(mail_elem + "\n")
                
f.close()
#移除临时目录
shutil.rmtree(temp_folder_path) 
