# # coding:utf-8
#
# from win32com import client as wc
# from docx import Document
#
# word = wc.Dispatch('Word.Application')
# doc = word.Documents.Open(u'C:/Users/Kurko/Desktop/用来测试的论文/需要检查的论文/硕士学位论文正文_1.doc')        # 目标路径下的文件
# doc.SaveAs(u'C:/Users/Kurko/Desktop/用来测试的论文/需要检查的论文/硕士学位论文正文_1.docx', 12, False, "", True, "", False, False, False, False)  # 转化后路径下的文件
# doc.Close()
# word.Quit()
#
# doc = Document('C:/Users/Kurko/Desktop/用来测试的论文/需要检查的论文/硕士学位论文正文_1.docx')
#
# pos1=None
# pos2=None
#
# for i in range(len(doc.paragraphs)):
#     print(doc.paragraphs[i].text.strip())
#     if doc.paragraphs[i].text.strip()=="参考文献":
#         pos1=i
#     if pos1 is not None and doc.paragraphs[i].text.strip()== "":
#         pos2=i
#         break
# print(pos1, pos2)
# for i in range(pos1, pos2+1):
#     print(i, doc.paragraphs[i].text)
# literList=[]

from win32com import client as wc
import re
import sys


def preproccess_file(file):
    '''文件预处理'''
    # 对文件内容预处理(把文本集中到一个字符串中）

    para_list = []
    for idx, para in enumerate(file.paragraphs):
        print(para.text)
        para_list.append(para)
    return para_list


def extract(keyword, para_list):
    '''按关键词提取内容'''
    # 使用正则提取关键字后面的数字
    is_reference = False
    refer_list = []

    for idx, para in enumerate(para_list):
        if not is_reference:
            result = re.match('{}'.format(keyword), para.text)
            if result:
                is_reference = True
        elif is_reference:
            if para.text.strip() == "":
                break
            print(para.text)
            if para.text.strip()[0] != '[':
                curLen = len(refer_list)
                refer_list[curLen - 1] += para.text
            else:
                refer_list.append(para.text)
    for i in range(len(refer_list)):
        print(i, refer_list[i])

    return refer_list


def check(refer_list):
    print("对参考文件进行格式检查")
    # 对引用部分提取匹配
    info_list = []
    # 提取基本格式，并检查
    for para in refer_list:
        # print statistical data;
        match = re.match('^\[(\d+)\].*', para)
        # print(para.text)
        if match:
            res = re.findall('\[(\w)\]', para)  # 提取[]中数字和文本类别
            info_list.append(res)
        else:
            print("引用开头格式有误，内容为：", para)

    # 次序检查，期刊类别检查
    old = 0
    for idx, info in enumerate(info_list):
        if int(info[0]) != old + 1:
            print("引用顺序有误:", refer_list[idx].text)
        old = int(info[0])
        if len(info) > 1:
            if (info[1] == 'J'):
                print("引用了期刊：", refer_list[idx].text)
        # TODO 其他检查

    # print(result)
    # return result


def main(path):
    print(path)
    real_path = path
    if path.endswith(".doc"):
        word = wc.Dispatch('Word.Application')
        doc = word.Documents.Open(path)
        doc.SaveAs(path + "x", 12, False, "", True, "", False, False, False, False)
        doc.Close()
        word.Quit()
        real_path = path + "x"
    elif not path.endswith(".docx"):
        print("请检查文件后缀名是否有效！")
        return

    keyword = '参考文献'
    word = wc.Dispatch('Word.Application')
    print(real_path)
    file = word.Documents.Open(path)
    file_text = preproccess_file(file)
    reference = extract(keyword, file_text)
    check(reference)


if __name__ == '__main__':
    if len(sys.argv) != 2:
        print("usage: python3 main.py <filepath>")
        exit(0)
    path = sys.argv[1]
    main(path)
