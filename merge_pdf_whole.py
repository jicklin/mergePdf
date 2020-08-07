# coding=utf-8
# 多个pdf合二为一
import easygui
import sys
import os
import re
import codecs
from PyPDF2 import PdfFileReader, PdfFileMerger, PdfFileWriter
import logging
import json

logging.basicConfig(level=logging.INFO,  # 控制台打印的日志级别
                    filename='merger_pdf.log',
                    filemode='a',  ##模式，有w和a，w就是写模式，每次都会重新写日志，覆盖之前的日志
                    # a是追加模式，默认如果不写的话，就是追加模式
                    format=
                    '%(asctime)s - %(pathname)s[line:%(lineno)d] - %(levelname)s: %(message)s'
                    # 日志格式
                    )


def traverse_pdf(pdf_obj, root_dir_path):
    for root, dirs, files in os.walk(root_dir_path):
        for file in files:
            file_full_path = os.path.join(root, file)
            file_ext = os.path.splitext(file)[1]
            # 只要pdf格式的
            if file_ext == '.pdf' or file_ext == '.PDF':
                """
                把不动产单元号揪出来 并且只要前边的19位
                """
                macth_obj = re.search(r'[A-Z0-9]+', file)
                unit_code = macth_obj.group()
                code = unit_code[0:19]
                if code not in pdf_obj:
                    file_list = [file_full_path]
                    pdf_obj[code] = file_list
                else:
                    pdf_obj[code].append(file_full_path)

    return pdf_obj


def traverse_pdf_2(pdf_obj, root_dir_path):
    pdf_obj = []
    for root, dirs, files in os.walk(root_dir_path):
        for file in files:
            file_full_path = os.path.join(root, file)
            file_ext = file.split('.')[-1].lower()

            # 只要pdf格式的
            if file_ext == 'pdf':
                pdf_obj.append(file_full_path)


    return pdf_obj


def format_pdf_list(root_dir_path):
    """
    逻辑出所有的pdf文件
    :param root_dir_path:
    :return:
    """
    pdf_obj = {}
    return traverse_pdf_2(pdf_obj, root_dir_path)


def merge_pdf(pdf_list, root_dir_path):
    """
    拼接pdf
    :param pdf_list:
    :param root_dir_path:
    :return:
    """

    count = 0
    error_list = []

    for key in pdf_list.keys():
        try:
            merger = PdfFileMerger()
            logging.info('总共需要处理%s个单元数据，当前是第%s个',len(pdf_list),count+1)

            logging.info('开始遍历单元数据%s->%s', key, json.dumps(pdf_list[key], ensure_ascii=False))

            for file_path in pdf_list[key]:

                f = codecs.open(file_path, 'rb')
                file_rd = PdfFileReader(f)
                if file_rd.isEncrypted:
                    logging.warn('不支持加密后的文件: %s', file_path)
                    continue
                merger.append(file_rd)
                logging.info('开始合并文件：%s', file_path)
                f.close()

            out_file_path = os.path.join(os.path.abspath(root_dir_path), key + ".pdf")
            merger.write(out_file_path)
            logging.info('单元：%s 合并后输出文件：%s',key, out_file_path)

            merger.close()
            count = count + 1
        except BaseException as e:
            error_list.append(key)
            logging.error('尝试合并文件错误,单元为：%s', key, exc_info=True)
            easygui.exceptionbox()
            pass
    logging.info('恭喜马佳佳同学合并成功，共成功合并%s个单元,失败的单元如下：%s',count,json.dumps(error_list))

    easygui.msgbox('恭喜马佳佳同学合并成功，共成功合并{}个单元,失败的单元如下：{}'.format(count,json.dumps(error_list)) )


    pass


def merge_pdf_2(pdf_list, root_dir_path):
    merger = PdfFileMerger()

    logging.info('开始遍历单元数据%s', json.dumps(pdf_list, ensure_ascii=False))

    for file_path in pdf_list:

        f = codecs.open(file_path, 'rb')
        file_rd = PdfFileReader(f)
        if file_rd.isEncrypted:
            logging.warn('不支持加密后的文件: %s', file_path)
            continue
        merger.append(file_rd)
        logging.info('开始合并文件：%s', file_path)
        f.close()

    out_file_path = os.path.join(os.path.abspath(root_dir_path),   "合并后的文件.pdf")
    merger.write(out_file_path)
    logging.info('合并后输出文件：%s', out_file_path)
    merger.close()
    easygui.msgbox('恭喜马佳佳同学合并成功，成功合并')



if __name__ == '__main__':

    '''
    选择所有pdf所在的文件的根目录
    '''
    root_dir_path = easygui.diropenbox(msg='选择根文件夹', title='浏览文件夹')

    if root_dir_path is None:
        easygui.msgbox('未选择根目录没法处理哦', title='提示')
        sys.exit()

    logging.info('选择的根目录是 %s ', root_dir_path)

    '''
    1.遍历及整理名字类似的PDF文件
    构成类似
    {
    "320623100001JC01001":["D:/地籍调查表320623100001JC01001XXX.pdf","D:/房屋320623100001JC01001XXX.pdf"]

    }
    '''
    logging.info('开始整理PDF文件')
    pdf_list = format_pdf_list(root_dir_path)

    if pdf_list is None:
        easygui.msgbox('警告', '未找到符合条件的PDF')
        sys.exit()
    logging.info('整理之后的不动产单元下的PDF文件是 %s ', json.dumps(pdf_list, ensure_ascii=False))

    out_path_str = root_dir_path + '//out'
    if not os.path.exists(out_path_str):
        os.makedirs(out_path_str)

    '''
    遍历组织好的拼接一下
    '''
    # merge_pdf(pdf_list, root_dir_path + '//out')
    merge_pdf_2(pdf_list, root_dir_path + '//out')
