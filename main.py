from win32com import client as wc
import re
import sys
from docx2python import docx2python


def extract(keyword, para_list):
    '''
    按关键词提取内容
    :param keyword: 字符串
    :param para_list: 应匹配的母串
    :return: 返回匹配成功的位置之后的所有段落的list
    '''
    # 使用正则提取关键字后面的数字
    is_reference = False
    refer_list = []

    for idx, para in enumerate(para_list):
        if not is_reference:
            result = re.match('{}'.format(keyword), para.strip())
            if result and len(para.strip()) == 4:
                is_reference = True
        elif is_reference:
            if para.strip() == "" or para.strip() == "作者简历":
                break
            if not (para.strip()[0] == '[' or '0' <= para.strip()[0] <= '9'):
                curLen = len(refer_list)
                refer_list[curLen - 1] += para
            else:
                refer_list.append(para)
    for i in range(len(refer_list)):
        print(refer_list[i])

    return refer_list


def _checkNum(refer_list):
    print("对参考文献篇数进行检查")
    print("篇数为：" + str(len(refer_list)))
    if len(refer_list) < 20:
        return False
    return True


def check(refer_list):
    """
    :param refer_list: 引用段落string格式的list
    :return:  直接输出结果，无返回
    """
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
            print("引用顺序有误:", refer_list[idx])
        old = int(info[0])
        if len(info) > 1:
            if (info[1] == 'J'):
                print("引用了期刊：", refer_list[idx])
        # TODO 其他检查

    # 篇数检查
    if not _checkNum(refer_list):
        print("参考文献篇数少于20篇！")



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
    file = docx2python(real_path)
    temp = file.text.split('\n')
    content = []
    for i in range(len(temp)):
        if i % 2 == 0:
            content.append(temp[i])

    # 可以输出reference看看，已经是text了
    reference = extract(keyword, content)
    check(reference)


if __name__ == '__main__':
    if len(sys.argv) != 2:
        print("usage: python3 main.py <filepath>")
        exit(0)
    path = sys.argv[1]
    main(path)
