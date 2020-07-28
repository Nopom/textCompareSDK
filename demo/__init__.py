import os
import zipfile
import difflib
import time
from xml.dom.minidom import parse
import json
from collections import OrderedDict
import shutil
import pdfplumber
import re

#文件后缀名修改   方便docx转成zip提取zip中的document.xml文件
def docx_zip(num, path):
    # num = 1  docx to zip
    # num = 2  zip to docx
    newtype = ''
    if num == 1:
        newtype = 'zip'
    elif num == 2:
        newtype = 'docx'
    path_list = path.split('.')
    del path_list[-1]
    new_path = ''
    for p in path_list:
        new_path += p + '.'
    new_path += newtype
    os.rename(path, new_path)
    return new_path

 #获取document.xml中的段落信息和表格信息
def document_work(rootdata, file):

    document_xml_con = []
    document_xml_index = {}
    table_xml_con = {}
    table_xml_doc = {}
    table_xml_index = {}
    table_xml_tc = {}
    bodydata = rootdata.getElementsByTagName('w:body')
    '''
    xml中 为了获取段落内容标签  会获取起内容，但有些标签是不都存在的
    补充标签，但不给值，否则会出现 无此标签或者没人标签内容的错误
    '''
    for body in bodydata:
        for p in body.getElementsByTagName('w:p'):
            if len(p.getElementsByTagName('w:t')) == 0:
                data = file.createTextNode(' ')
                node = file.createElement('w:r')
                node2_1 = file.createElement('w:rPr')
                node2_2 = file.createElement('w:t')
                node2_2.appendChild(data)
                node.appendChild(node2_1)
                node.appendChild(node2_2)
                p.appendChild(node)
            for r in p.getElementsByTagName('w:r'):
                names = []
                for rs in r.childNodes:
                    name = rs.nodeName
                    names.append(name)
                if 'w:t' in names:
                    if 'w:rPr' not in names:
                        child_t = r.getElementsByTagName('w:t')[0]
                        clone_t = child_t.cloneNode(True)
                        node_rpr = file.createElement('w:rPr')
                        r.appendChild(node_rpr)
                        r.appendChild(clone_t)
                        r.removeChild(child_t)
        # 段落
        #获取全部段落内容，段落为一个单位
        num_p = 0
        for nodes in body.childNodes:
            if nodes.nodeName == 'w:p':
                dict_index_index = {}
                doc_con = ''
                num_r = 0
                for doc in nodes.childNodes:
                    if doc.nodeName == 'w:r':
                        for t in doc.getElementsByTagName('w:t'):
                            doc_con += t.firstChild.data
                            dict_index_index[num_r] = t.firstChild.data
                            document_xml_index[num_p] = dict_index_index
                    num_r += 1
                num_p += 1
                document_xml_con.append(doc_con)
        # 表格
        num_tbl = 0
        for nodes in body.childNodes:
            if nodes.nodeName == 'w:tbl':
                num_tr = 0
                dict_tr = {}
                str_tr = ''
                tr_list = []
                list_tr = {}
                for ts in nodes.childNodes:
                    if ts.nodeName == 'w:tr':
                        num_tc = 0
                        dict_tc = {}
                        str_tc = ''
                        list_tc = []
                        for tc in ts.childNodes:
                            if tc.nodeName == 'w:tc':
                                num_p = 0
                                dict_p = {}
                                list_tp = []
                                for tp in tc.childNodes:
                                    if tp.nodeName == 'w:p':
                                        num_r = 0
                                        dict_r = {}
                                        str_tt = ''
                                        for tr in tp.childNodes:
                                            if tr.nodeName == 'w:r':
                                                num_tt = 0
                                                dict_tt = {}
                                                for tt in tr.childNodes:
                                                    if tt.nodeName == 'w:t':
                                                        dict_tt[num_tt] = tt.firstChild.data
                                                        dict_r[num_r] = dict_tt
                                                        dict_p[num_p] = dict_r
                                                        dict_tc[num_tc] = dict_p
                                                        dict_tr[num_tr] = dict_tc
                                                        table_xml_index[num_tbl] = dict_tr
                                                        str_tc += tt.firstChild.data
                                                        str_tt += tt.firstChild.data
                                                    num_tt += 1
                                            num_r += 1
                                        list_tp.append(str_tt)
                                    num_p += 1
                                list_tc.append(list_tp)
                            num_tc += 1
                        str_tr += str_tc
                        tr_list.append(str_tc)
                        list_tr[num_tr] = list_tc
                    num_tr += 1
                table_xml_con[num_tbl] = str_tr
                table_xml_doc[num_tbl] = tr_list
                table_xml_tc[num_tbl] = list_tr
                num_tbl += 1
    # print(document_xml_con)
    # print(table_xml_con)
    # print(document_xml_index)
    return document_xml_con

#将第一层对比文档内容分开
def cdl2diff(cdl):
    '''
    使用有序的dict 避免段落出现误差，使得最后结果不一致
    :param cdl:
    :return:
    '''
    diff_doc1 = OrderedDict()
    diff_doc2 = OrderedDict()
    for (i, v), n in zip(cdl.items(), range(len(cdl))):
        if n % 2 == 0:
            diff_doc1[i] = v
        else:
            diff_doc2[i] = v
    return diff_doc1,diff_doc2

#区分差异类型
def diff2cdl(diff_words):
    '''
    根据diff_words的输出内容，加入规则判断不一样的修改内容
    若后面有内容没有识别出来，可能是这里的判断逻辑的遗漏
    建议输出diff_words查看，并在最后输出一下相应修改类型的内容，查看是否一致

    '''
    #diff 之后，不一样的区别类型分类
    # for i in diff_words:
    #     print(i)
    index_list = []
    cdl_add = OrderedDict()
    for i in range(len(diff_words)):
        if '?' == diff_words[i][0] and '+' in diff_words[i]:
            if diff_words[i - 2][0] == '?':
                cdl_add[i - 3] = diff_words[i - 3]
                index_list.append(i - 3)
            else:
                cdl_add[i - 2] = diff_words[i - 2]
                index_list.append(i - 2)
            cdl_add[i - 1] = diff_words[i - 1]
            index_list.append(i - 1)
    cdl_delete = OrderedDict()
    for i in range(len(diff_words)):
        if '?' == diff_words[i][0] and '-' in diff_words[i]:
            # print(diff_words[i-1])
            # print(diff_words[i+1])
            cdl_delete[i - 1] = diff_words[i - 1]
            cdl_delete[i + 1] = diff_words[i + 1]
            index_list.append(i - 1)
            index_list.append(i + 1)
    cdl_update = OrderedDict()
    for i in range(len(diff_words)):
        if '?' == diff_words[i][0] and '^' in diff_words[i]:
            cdl_update[i - 1] = diff_words[i - 1]
            index_list.append(i - 1)
    cdl_other = OrderedDict()
    for i in range(len(diff_words)):
        if i not in index_list and diff_words[i][0] != '?':
            cdl_other[i] = diff_words[i]
    return cdl_delete,cdl_add,cdl_update,cdl_other

#正则化数据类型
def re_time(m):
    '''
    将修改内容为数字或者时间类型的数据按数据和时间格式输出
    :param m:
    :return:
    '''
    pattern_num = re.compile(r'\d{2,}|\d+.\d+|'
                             r'[\u4e00\u4e8c\u4e09\u56db\u4e94\u516d\u4e03\u516b\u4e5d\u5341]{2,}')
    num_search = re.search(pattern_num, m)
    index_num_list = []
    if num_search != None:
        end_flag = num_search.end()
        end_flag_1 = num_search.end()
        start_flag_1 = num_search.start()
        index_num_list.append([start_flag_1, end_flag_1])
        flag = True
        m_f = m
        while flag:
            m_f = m_f[end_flag:]
            time_search_flag = re.search(pattern_num, m_f)
            if time_search_flag != None:
                end_flag = time_search_flag.end()
                start_flag = time_search_flag.start()
                start_flag_1 += end_flag
                end_flag_1 += end_flag
                start_flag_1 = end_flag_1 - (end_flag - start_flag)
                index_num_list.append([start_flag_1, end_flag_1])
            else:
                flag = False
    # print('整数：',index_num_list)
    pattern_time = re.compile(r"\d{2}/\d{2}/\d{4}\b|\d{1}:\d{2}|"
                              r"\d{2}:\d{2}|\d{4}-\d{1,2}-\d{1,2}\s\d{1,2}:\d{1,2}|"
                              r"\d{4}-\d{1,2}-\d{1,2}\s\d{1,2}:\d{1,2}|"
                              r"\d{2,4}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日")  # 定义匹配模式
    time_search = re.search(pattern_time, m)
    index_time_list = []
    if time_search != None:

        end_flag = time_search.end()
        end_flag_1 = time_search.end()
        start_flag_1 = time_search.start()
        index_time_list.append([start_flag_1, end_flag_1])

        flag = True
        m_f = m
        while flag:
            m_f = m_f[end_flag:]
            time_search_flag = re.search(pattern_time, m_f)
            if time_search_flag != None:
                end_flag = time_search_flag.end()
                start_flag = time_search_flag.start()
                end_flag_1 += end_flag
                start_flag_1 += end_flag_1 - (end_flag - start_flag)
                index_time_list.append([start_flag_1, end_flag_1])
            else:
                flag = False
    index_num_list_copy = index_num_list.copy()
    for time in index_time_list:
        for num in index_num_list_copy:
            if num[0] >= time[0] and num[1] <= time[1]:
                index_num_list.remove(num)
    for num in index_num_list:
        index_time_list.append(num)
    return index_time_list

#去除相同类型修改数据（现在已经停用）
def dict_remove_duplication(dict1, dict2):

    dict_flag = {}
    num_flag = 0
    for (i1, v1), (i2, v2) in zip(dict1.items(),
                                  dict2.items()):

        for v1_1, v2_1 in zip(v1, v2):
            dict_flag[num_flag] = str(v1_1) + '&&' + str(v2_1)
            num_flag += 1
        num_flag += 1
    func = lambda z: dict([(x, y) for y, x in z.items()])  # 字典键值对位置互换
    result_func = func(func(dict_flag))
    func2 = lambda z: dict([(x, str(y).split('&&')) for y, x in z.items()])  # 字典键值对位置互换
    result_func2 = func2(func(result_func))
    return result_func2

def con_duplication(i1, i2, m, index_list):
    '''
       筛选与要求符合的文字
       '''
    num_con = []
    dict_con = {}
    if len(index_list) == 0:
        num_con.append(m[i1 + 2:i2 + 2])
    else:
        for in_ls in index_list:
            if i1 + 2 >= in_ls[0] and i2 + 2 <= in_ls[1]:
                print('=========', m[i1 + 2:i2 + 2])
                str_flag = i1 + i2
                dict_con[str_flag] = m[in_ls[0]:in_ls[1]]
            else:
                print('---------', m[i1 + 2:i2 + 2])
                num_con.append(m[i1 + 2:i2 + 2])
    # print(num_con)
    # print('==', dict_con)
    for _, v in dict_con.items():
        num_con.append(v)
    return num_con

def xml_extract(originalPath,comparePath):

    # docx转zip  参数为1
    # zip转docx  参数为2
    originalPath_new = ''
    comparePath_new = ''
    path_list = originalPath.split('.')
    if path_list[-1] != 'zip':
        originalPath_new = docx_zip(1, originalPath)
    path_list = comparePath.split('.')
    if path_list[-1] != 'zip':
        comparePath_new = docx_zip(1, comparePath)
    # 提取对比文档
    file_list = [originalPath_new, comparePath_new]
    files = []
    # 获取zip中需要分析的document.xml文件
    for file in file_list:
        s = zipfile.ZipFile(file)
        f = s.open('word/document.xml', 'r')
        files.append(parse(f))
        s.close()

    return files, originalPath_new, comparePath_new

def pdf_extract(path):

    text_con = ''  # 存放文本内容
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            for ps in page.extract_text().split('\n'):
                for p in ps:
                    if p != ' ':
                        text_con += p
    return [text_con]

def doc_extract(files):

    file1 = files[0]
    file2 = files[1]
    #获取xml中的文本信息（con为文本信息，index为文本位置）
    document_xml_con1 = document_work(file1.documentElement, file1)

    document_xml_con2 = document_work(file2.documentElement, file2)

    return document_xml_con1, document_xml_con2

def document_compare_pdf(document_con1, document_con2):

    # 对比文档增加的文档内容和位置
    dict_file2_add_con = {}
    # 原始文档删除的文档内容和位置
    dict_file1_delete_con = {}
    # 原始文档修改的文档内容和位置
    dict_file1_update_con = {}
    # 对比文档修改的文档内容和位置
    dict_file2_update_con = {}
    diff = difflib.Differ()
    # 差异文本内容 详细对比
    content_diff = diff.compare(document_con1, document_con2)
    diff_words = list(content_diff)
    content_diff.close()
    cdl_delete, cdl_add, cdl_update, cdl_other = diff2cdl(diff_words)
    # 类型：对比文档中增加内容
    diff_doc1, diff_doc2 = cdl2diff(cdl_add)

    for (n1, m1), (n2, m2) in zip(diff_doc1.items(), diff_doc2.items()):
        ss = difflib.SequenceMatcher(lambda x: x == " ", m1[2:], m2[2:])
        num_add_all = []  # 增加内容位置
        num_add_con = []  # 增加内容文本
        for (tag, i1_1, i2_1, d1_1, d2_1) in ss.get_opcodes():
            # print(("%7s file1[%d:%d] (%s) file2[%d:%d] (%s)" %
            #        (tag, i1_1, i2_1, m1[i1_1 + 2:i2_1 + 2],
            #         d1_1, d2_1, m2[d1_1 + 2:d2_1 + 2])))
            if tag == 'insert':
                num_add = []
                for num in range(d1_1, d2_1):
                    num_add.append(num)
                num_add_all.append(num_add)
                num_add_con.append(m2[d1_1 + 2:d2_1 + 2])
                dict_file2_add_con[n2] = num_add_con

    # 类型：原始文档中删除的内容
    # 第二层 段落内部差异
    diff_doc1, diff_doc2 = cdl2diff(cdl_delete)
    for (n1, m1), (n2, m2) in zip(diff_doc1.items(), diff_doc2.items()):
        ss = difflib.SequenceMatcher(lambda x: x == " ", m1[2:], m2[2:])
        num_delete_all = []  # 删除内容位置
        num_delete_con = []  # 删除内容文本

        for (tag, i1_1, i2_1, d1_1, d2_1) in ss.get_opcodes():

            if tag == 'delete':
                # print(("%7s file1[%d:%d] (%s) file2[%d:%d] (%s)" %
                #        (tag, i1_1, i2_1, m1[i1_1 + 2:i2_1 + 2],
                #         d1_1, d2_1, m2[d1_1 + 2:d2_1 + 2])))
                num_delete = []
                for num in range(i1_1, i2_1):
                    num_delete.append(num)
                num_delete_all.append(num_delete)
                num_delete_con.append(m1[i1_1 + 2:i2_1 + 2])
                dict_file1_delete_con[n1] = num_delete_con

    # 类型：原始和对比文档中修改的内容
    # 第二层 段落内部差异
    diff_doc1, diff_doc2 = cdl2diff(cdl_update)
    for (n1, m1), (n2, m2) in zip(diff_doc1.items(), diff_doc2.items()):
        # 加入正则化项  搜索时间类型
        # 帅选符合要求字段下标
        index_list_m1 = re_time(m1)
        index_list_m2 = re_time(m2)
        # print(m1)
        # print(index_list_m1)
        ss = difflib.SequenceMatcher(lambda x: x == " ", m1[2:], m2[2:])
        num_update1_all = []
        num_update1_con = []
        num_update2_all = []
        num_update2_con = []
        for (tag, i1_1, i2_1, d1_1, d2_1) in ss.get_opcodes():
            if tag == 'replace':
                print(("%7s file1[%d:%d] (%s) file2[%d:%d] (%s)" %
                   (tag, i1_1, i2_1, m1[i1_1 + 2:i2_1 + 2],
                    d1_1, d2_1, m2[d1_1 + 2:d2_1 + 2])))
                num_update1 = []
                for num in range(i1_1, i2_1):
                    num_update1.append(num)
                num_update1_all.append(num_update1)
                # 筛选符合要求的字段
                # print(num_update1_con)
                num_update1_con = con_duplication(
                    num_update1_con, i1_1, i2_1, m1, index_list_m1)
                # print(num_update1_con)
                dict_file1_update_con[n1] = num_update1_con

                num_update2 = []
                for num in range(d1_1, d2_1):
                    num_update2.append(num)
                num_update2_all.append(num_update2)

                # 筛选字段
                num_update2_con = con_duplication(
                    num_update2_con, d1_1, d2_1, m2, index_list_m2)

                dict_file2_update_con[n2] = num_update2_con

    return dict_file1_delete_con, dict_file2_add_con, \
           dict_file1_update_con, dict_file2_update_con

def document_compare_docx(document_con1, document_con2):
    '''
    此项目核心函数，出现bug的大多数逻辑都在此函数
    此函数代码较多，但是逻辑基本一致，出问题之后
    1.输出此函数的主要对象，确定错误大体位置
    2.从第二层对比开始判断错误位置
    3.从第一层输出判断错误位置
    4.里面涉及到的工具类里的方法很重要，错误关键所在地
    :return:
    '''

    #对比文档增加的文档内容和位置
    dict_file2_add_index = {}
    dict_file2_add_con = {}
    # 原始文档删除的文档内容和位置
    dict_file1_delete_index = {}
    dict_file1_delete_con = {}
    # 原始文档修改的文档内容和位置
    dict_file1_update_index = {}
    dict_file1_update_con = {}
    # 对比文档修改的文档内容和位置
    dict_file2_update_index = {}
    dict_file2_update_con = {}
    #获取两个文档的全部文本
    #全部文本比较
    s = difflib.SequenceMatcher(lambda x: x == " ",
                                document_con1, document_con2)
    diff = difflib.Differ()
    #第一层比较  文本比较（段落为单位）
    for (tag, i1, i2, d1, d2) in s.get_opcodes():

        # print(("%7s \n file1[%d:%d] (%s) \n file2[%d:%d] (%s)" %
        #        (tag, i1, i2, document_con1[i1:i2], d1, d2,
        #         document_con2[d1:d2])))

        #tag差异类型  i 原始文本差异位置   d 对比文本差异位置
        #获取段落有差异的内容 且段落个数大于一段，1、进行段落匹配 2、段落详细比较
        if tag == 'replace' and (len(document_con1[i1:i2]) > 1
                                 or len(document_con2[d1:d2]) > 1) :

            #差异文本内容 详细对比
            content_diff = diff.compare(document_con1[i1:i2], document_con2[d1:d2])
            diff_words = list(content_diff)
            content_diff.close()
            cdl_all = {}
            for i in range(len(diff_words)):
                if '-' == diff_words[i][0] or '+' == diff_words[i][0]:
                    cdl_all[i] = diff_words[i]

            #根据compare输出结果  分出修改类型
            '''
            此处得出的不同类型的修改内容  会影响后续的第二层文本对比
            '''
            cdl_delete, cdl_add, cdl_update, cdl_other = diff2cdl(diff_words)
            #筛选错落或缺失的文本内容
            cdl_file1 = OrderedDict()
            cdl_file1_con = []
            cdl_file2 = OrderedDict()
            cdl_file2_con = []
            for i,v in cdl_all.items():
                if v[0] == '-':
                    cdl_file1[i] = v
                    cdl_file1_con.append(v[2:])
                else:
                    cdl_file2[i] = v
                    cdl_file2_con.append(v[2:])

            doc1 = []
            doc1_index = []
            for e in range(i1, i2):
                doc1.append(document_con1[e])
                doc1_index.append(e)

            #筛选出缺失或错位的段落
            s_file1 = difflib.SequenceMatcher(lambda x: x == " ", cdl_file1_con,doc1)
            for (tag, i1_1, i2_1, d1_1, d2_1) in s_file1.get_opcodes():
                if tag == 'insert':
                    for i in range(len(doc1)):
                        if i == i1_1:
                            r = doc1_index[i]
                            doc1_index.remove(r)
            doc2 = []
            doc2_index = []
            for e in range(d1, d2):
                doc2.append(document_con2[e])
                doc2_index.append(e)

            s_file2 = difflib.SequenceMatcher(lambda x: x == " ", cdl_file2_con, doc2)
            for (tag, i1_1, i2_1, d1_1, d2_1) in s_file2.get_opcodes():
                if tag == 'insert':
                    for i in range(len(doc2)):
                        if i == i1_1:
                            r = doc2_index[i]
                            doc2_index.remove(r)
            #重新进行位置排列之后的前后对应关系
            flag_file1 = OrderedDict()
            for i, e in zip(cdl_file1,doc1_index ):
                flag_file1[i] = e

            flag_file2 = OrderedDict()
            for i, e in zip(cdl_file2, doc2_index):
                flag_file2[i] = e


            '''
            从这里开始 进入第二层文本对比，得到的就是最终的结果，
            如果后续输出存在细微差异 建议从这找问题
            这段代码判断修改种类较多 
            '''
            # 类型：对比文档中增加内容
            # 第二层 段落内部差异（句子为单位）
            diff_doc1,diff_doc2 = cdl2diff(cdl_add)

            for (n1,m1),(n2,m2) in zip(diff_doc1.items(),diff_doc2.items()):
                ss = difflib.SequenceMatcher(lambda x: x == " ",m1[2:], m2[2:])
                num_add_all = [] #增加内容位置
                num_add_con = [] #增加内容文本
                for (tag, i1_1, i2_1, d1_1, d2_1) in ss.get_opcodes():
                    # print(("%7s file1[%d:%d] (%s) file2[%d:%d] (%s)" %
                    #        (tag, i1_1, i2_1, m1[i1_1 + 2:i2_1 + 2],
                    #         d1_1, d2_1, m2[d1_1 + 2:d2_1 + 2])))
                    if tag == 'insert':
                        num_add = []
                        for num in range(d1_1, d2_1):
                            num_add.append(num)
                        num_add_all.append(num_add)
                        num_add_con.append(m2[d1_1 + 2:d2_1 + 2])
                        for index, value in flag_file2.items():
                            if n2 == index:
                                dict_file2_add_con[value] = num_add_con
                                dict_file2_add_index[value] = num_add_all

            # 类型：原始文档中删除的内容
            # 第二层 段落内部差异
            diff_doc1, diff_doc2 = cdl2diff(cdl_delete)
            for (n1,m1),(n2,m2) in zip(diff_doc1.items(),diff_doc2.items()):
                ss = difflib.SequenceMatcher(lambda x: x == " ", m1[2:], m2[2:])
                num_delete_all = []#删除内容位置
                num_delete_con = []#删除内容文本

                for (tag, i1_1, i2_1, d1_1, d2_1) in ss.get_opcodes():

                    if tag == 'delete':
                        # print(("%7s file1[%d:%d] (%s) file2[%d:%d] (%s)" %
                        #        (tag, i1_1, i2_1, m1[i1_1 + 2:i2_1 + 2],
                        #         d1_1, d2_1, m2[d1_1 + 2:d2_1 + 2])))
                        num_delete = []
                        for num in range(i1_1, i2_1):
                            num_delete.append(num)
                        num_delete_all.append(num_delete)
                        num_delete_con.append(m1[i1_1 + 2:i2_1 + 2])
                        for index, value in flag_file1.items():
                            if n1 == index:
                                dict_file1_delete_con[value] = num_delete_con
                                dict_file1_delete_index[value] = num_delete_all

            # 类型：原始和对比文档中修改的内容
            # 第二层 段落内部差异
            diff_doc1, diff_doc2 = cdl2diff(cdl_update)
            for (n1, m1),(n2, m2) in zip(diff_doc1.items(), diff_doc2.items()):
                #加入正则化项  搜索时间类型
                # 帅选符合要求字段下标
                index_list_m1 = re_time(m1)
                index_list_m2 = re_time(m2)
                ss = difflib.SequenceMatcher(lambda x: x == " ", m1[2:], m2[2:])
                num_update1_all = []
                num_update1_con = []
                num_update2_all = []
                num_update2_con = []
                for (tag, i1_1, i2_1, d1_1, d2_1) in ss.get_opcodes():
                    if tag == 'replace':
                        print(("%7s file1[%d:%d] (%s) file2[%d:%d] (%s)" %
                           (tag, i1_1, i2_1, m1[i1_1 + 2:i2_1 + 2],
                            d1_1, d2_1, m2[d1_1 + 2:d2_1 + 2])))

                        num_update1 = []
                        for num in range(i1_1, i2_1):
                            num_update1.append(num)
                        num_update1_all.append(num_update1)
                        # 筛选符合要求的字段
                        num_update1_con = con_duplication(
                                        i1_1,i2_1,m1,index_list_m1)
                        # print(num_update1_con)
                        for index, value in flag_file1.items():
                            if n1 == index:
                                dict_file1_update_con[value] = num_update1_con
                                dict_file1_update_index[value] = num_update1_all

                        num_update2 = []
                        for num in range(d1_1, d2_1):
                            num_update2.append(num)
                        num_update2_all.append(num_update2)

                        # 筛选字段
                        num_update2_con = con_duplication(
                            d1_1, d2_1, m2, index_list_m2)
                        for index, value in flag_file2.items():
                            if n2 == index:
                                dict_file2_update_con[value] = num_update2_con
                                dict_file2_update_index[value] = num_update2_all

            for n,m in cdl_other.items():
                if m[0] == '+':
                    num_add_all = []
                    num_add = []
                    for l in range(len(m[2:])):
                        num_add.append(l)
                    num_add_all.append(num_add)
                    for index, value in flag_file2.items():
                        if n == index:
                            dict_file2_add_con[value] = m[2:]
                            dict_file2_add_index[value] = num_add_all
                else:
                    num_delete_all = []
                    num_delete = []
                    for l in range(len(m[2:])):
                        num_delete.append(l)
                    num_delete_all.append(num_delete)
                    for index, value in flag_file1.items():
                        if n == index:
                            dict_file1_delete_con[value] = m[2:]
                            dict_file1_delete_index[value] = num_delete_all

        #获取有差异的段落 且一一匹配 直接进行第二层对比
        elif tag == 'replace' and len(document_con1[i1:i2]) == 1 and \
                                    len(document_con2[d1:d2]) == 1:
            # print(("%7s \n file1[%d:%d] (%s) \n file2[%d:%d] (%s)" %
            #        (tag, i1, i2, document_con1[i1:i2], d1, d2,
            #         document_con2[d1:d2])))
            index_list_m1 = re_time(document_con1[i1])
            index_list_m2 = re_time(document_con2[d1])
            ss = difflib.SequenceMatcher(lambda x: x == " ",
                                         document_con1[i1], document_con2[d1])

            num_add_all = []
            num_add_con = []
            num_delete_all = []
            num_delete_con = []
            num_update1_all = []
            num_update1_con = []
            num_update2_all = []
            num_update2_con = []

            for (tag, i1_1, i2_1, d1_1, d2_1) in ss.get_opcodes():
                # print(("%7s file1[%d:%d] (%s) file2[%d:%d] (%s)" %
                #       (tag, i1_1, i2_1, document_con1[i1][i1_1:i2_1],
                #         d1_1, d2_1,document_con2[d1][d1_1:d2_1])))
                if tag == 'insert':
                    num_add = []
                    for num in range(d1_1, d2_1):
                        num_add.append(num)
                    num_add_all.append(num_add)
                    num_add_con.append(document_con2[d1][d1_1:d2_1])
                    dict_file2_add_con[d1] = num_add_con
                    dict_file2_add_index[d1] = num_add_all
                elif tag == 'delete':
                    num_delete = []
                    for num in range(i1_1, i2_1):
                        num_delete.append(num)
                    num_delete_all.append(num_delete)
                    num_delete_con.append(document_con1[i1][i1_1:i2_1])
                    dict_file1_delete_con[i1] = num_delete_con
                    dict_file1_delete_index[i1] = num_delete_all
                elif tag == 'replace':
                    # print(("%7s file1[%d:%d] (%s) file2[%d:%d] (%s)" %
                    #                (tag, i1_1, i2_1, document_con1[i1][i1_1:i2_1],
                    #         d1_1, d2_1,document_con2[d1][d1_1:d2_1])))
                    num_update1 = []
                    for num in range(i1_1, i2_1):
                        num_update1.append(num)
                    num_update1_all.append(num_update1)
                    # 筛选字段
                    m1 = '@@' + document_con1[i1]
                    num_update1_con = con_duplication(
                        num_update1_con, i1_1, i2_1, m1, index_list_m1)
                    dict_file1_update_con[i1] = num_update1_con
                    dict_file1_update_index[i1] = num_update1_all


                    num_update2 = []
                    for num in range(d1_1, d2_1):
                        num_update2.append(num)
                    num_update2_all.append(num_update2)

                    #筛选字段
                    m2 = '@@' + document_con2[d1]
                    num_update2_con = con_duplication(
                        num_update2_con, d1_1, d2_1, m2, index_list_m2)
                    dict_file2_update_con[d1] = num_update2_con
                    dict_file2_update_index[d1] = num_update2_all
        #对比文档中增加的段落
        elif tag == 'insert':
            # print(("%7s \n file1[%d:%d] (%s) \n file2[%d:%d] (%s)" %
            #        (tag, i1, i2, document_con1[i1:i2], d1, d2,
            #         document_con2[d1:d2])))
            for ls in range(d1, d2):
                num_add_all = []
                num_add = []
                for l in range(len(document_con2[ls])):
                    num_add.append(l)
                num_add_all.append(num_add)
                dict_file2_add_con[ls] = [document_con2[ls]]
                dict_file2_add_index[ls] = num_add_all
        #原始文档中删除的段落
        elif tag == 'delete':
            # print(("%7s \n file1[%d:%d] (%s) \n file2[%d:%d] (%s)" %
            #        (tag, i1, i2, document_con1[i1:i2], d1, d2,
            #         document_con2[d1:d2])))
            for ls in range(i1,i2):
                num_delete_all = []
                num_delete = []
                for l in range(len(document_con1[ls])):
                    num_delete.append(l)
                num_delete_all.append(num_delete)
                dict_file1_delete_con[ls] = [document_con1[ls]]
                dict_file1_delete_index[ls] = num_delete_all

    return dict_file1_delete_con, dict_file2_add_con, \
        dict_file1_update_con, dict_file2_update_con

def dump_json(dict_file1_delete_con, dict_file2_add_con,
            dict_file1_update_con, dict_file2_update_con):
    '''
    最后一步写入json
    如果以上都未发现错误信息，那么最后输出不匹配  问题就在此处
    逻辑还需慢慢优化
    '''
    file_json = {'Replace': [], 'Original_document_delete': [], 'Contrast_document_add': []}

    #json加入修改内容
    num_flag = 0
    for (i1, v1), (i2, v2) in zip(dict_file1_update_con.items(),
                                  dict_file2_update_con.items()):
        if len(v1) == 1 or len(v2) == 1:
            file_json['Replace'].append({'originalItem_' + str(i1) + '_' + str(num_flag): v1[0],
                                         'modifyItem_' + str(i2) + '_' + str(num_flag): v2[0]})
            num_flag += 1
        if len(v1) > 1 or len(v2) > 1:
            for v11, v22 in zip(v1, v2):
                file_json['Replace'].append({'originalItem_' + str(i1) + '_' + str(num_flag): v11,
                                             'modifyItem_' + str(i2) + '_' + str(num_flag): v22})
                num_flag += 1
    num_flag = 0
    #json加入原始文档删除内容
    for i, con in dict_file1_delete_con.items():
        for c in con:
            file_json['Original_document_delete'].\
                append({'originalItem_' + str(i) + '_' + str(num_flag): c,
                        'modifyItem_' + str(i) + '_' + str(num_flag): ''})
            num_flag += 1

    #json加入对比文档增加内容
    num_flag = 0
    for i, con in dict_file2_add_con.items():
        for c in con:
            file_json['Contrast_document_add'].\
                append({'originalItem_' + str(i) + '_' + str(num_flag): '',
                        'modifyItem_' + str(i) + '_' + str(num_flag): c})
            num_flag += 1

    path = os.getcwd()
    if 'compareResult' not in os.listdir(path):
        os.mkdir(path + '\\' + 'compareResult')
    com_time = time.strftime("%Y%m%d%H%M%S", time.localtime(time.time()))
    with open(path + '\\' + 'compareResult\\' + com_time+'.json', 'w', encoding='utf-8') as f:
        json.dump(file_json, f, ensure_ascii=False, sort_keys=True, indent=4)
        print('对比结果JSON文件写入完成...')
    f.close()
    return com_time

def compare_docx(originalPath, comparePath):

    original = originalPath.split('.')[-1]
    compare = comparePath.split('.')[-1]
    dict_file1_delete_con = {}
    dict_file2_add_con = {}
    dict_file1_update_con = {}
    dict_file2_update_con = {}
    if original == 'docx' and compare == 'docx':
        files, originalPath_new, comparePath_new = xml_extract(originalPath, comparePath)
        document_xml_con1, document_xml_con2 = doc_extract(files)
        _ = docx_zip(2, originalPath_new)
        _ = docx_zip(2, comparePath_new)
        dict_file1_delete_con, dict_file2_add_con, \
        dict_file1_update_con, dict_file2_update_con = \
                    document_compare_docx(document_xml_con1, document_xml_con2)

    elif original == 'pdf' and compare == 'pdf':
        document_xml_con1 = pdf_extract(originalPath)
        document_xml_con2 = pdf_extract(comparePath)
        dict_file1_delete_con, dict_file2_add_con, \
        dict_file1_update_con, dict_file2_update_con = \
                    document_compare_pdf(document_xml_con1, document_xml_con2)

    com_time = dump_json(dict_file1_delete_con, dict_file2_add_con,
                        dict_file1_update_con, dict_file2_update_con)

    path = os.getcwd()
    if 'historyCompareDocx' not in os.listdir(path):
        os.mkdir(path + '\\' + 'historyCompareDocx')

    original = originalPath.split('\\')[-1]
    compare = comparePath.split('\\')[-1]
    # 将对比完文档复制到历史对比文档目录中文档
    os.mkdir(path + '\\' + 'historyCompareDocx\\' + com_time)

    shutil.copy(originalPath,
                path + '\\' + 'historyCompareDocx\\' + com_time + '\\' + original )
    shutil.copy(comparePath,
                path + '\\' + 'historyCompareDocx\\' + com_time + '\\' + compare )

if __name__ == '__main__':

    originalPath = r'D:\textCompare\原始文档0618.docx'
    # originalPath = r'D:\textCompare\test.pdf'
    comparePath= r'D:\textCompare\对比文档0618.docx'
    # comparePath= r'D:\textCompare\test-modify.pdf'
    compare_docx(originalPath, comparePath)
