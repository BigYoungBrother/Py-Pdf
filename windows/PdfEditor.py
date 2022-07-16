#!/usr/bin/env python
# -*- encoding: utf-8 -*-
"""
@desc    :   PDF编辑器
@author  :   Young
@date    :   2022-07-16 12:30
"""
import os

import fitz
from PyPDF2 import PdfFileWriter, PdfFileMerger, PdfFileReader


class PdfEditor:
    """
    提取某一页PDF
    """

    def get_page(self, from_pdf_file, page_num=None):
        pdf_reader = PdfFileReader(open(from_pdf_file, 'rb'))
        if page_num is None or page_num < 0:
            for i in range(pdf_reader.getNumPages()):
                self.do_get_page(from_pdf_file, i, pdf_reader)
            return
        if page_num < 1:
            print("待提取页码必须大于或等于1, 您的输入为: %s" % page_num)
            return
        if page_num > pdf_reader.getNumPages():
            print("待提取的页码不能超过PDF总页数, 您的输入为: %s, PDF总页数为: %s" % (page_num, pdf_reader.getNumPages()))
            return
        self.do_get_page(from_pdf_file, page_num - 1, pdf_reader)

    """
    获取pdf某一页
    """

    def do_get_page(self, from_pdf_file, page_num, pdf_reader):
        to_pdf_dir = os.path.dirname(os.path.abspath(from_pdf_file)) + '/split_pdf'
        if not os.path.exists(to_pdf_dir):
            # 判断文件夹是否存在, 文件夹不存在就创建
            os.makedirs(to_pdf_dir)
        to_pdf_file = to_pdf_dir + '/pdf_%s.pdf' % page_num
        pdf_writer = PdfFileWriter()
        pdf_writer.addPage(pdf_reader.getPage(page_num))
        output = open(to_pdf_file, 'wb')
        pdf_writer.write(output)
        output.close()

    """
    PDF转图片
    """

    def pdf_to_image(self, from_pdf_file):
        image_path = os.path.dirname(os.path.abspath(from_pdf_file)) + "/pdf_images"
        if not os.path.exists(image_path):
            # 判断文件夹是否存在, 文件夹不存在就创建
            os.makedirs(image_path)
        pdf_doc = fitz.open(from_pdf_file)
        for pg in range(pdf_doc.page_count):
            page = pdf_doc[pg]
            rotate = int(0)
            # 每个尺寸的缩放系数为1.3，这将为我们生成分辨率提高2.6的图像。
            # 此处若是不做设置，默认图片大小为：792X612, dpi=72
            # (1.33333333-->1056x816)   (2-->1584x1224)
            zoom_x = 1.33333333
            zoom_y = 1.33333333
            mat = fitz.Matrix(zoom_x, zoom_y).prerotate(rotate)
            pix = page.get_pixmap(matrix=mat, alpha=False)
            # 将图片写入指定的文件夹内
            pix.save(image_path + '/' + 'pdf_images_%s.png' % pg)

    """
    PDF合并
    @:param target_path: 源PDF文件目录
    @:param out_pdf_file: 输出PDF
    """

    def merge_pdf(self, target_path):
        pdf_lst = [f for f in os.listdir(target_path) if f.endswith('.pdf')]
        pdf_lst.sort()
        pdf_lst = [os.path.join(target_path, filename) for filename in pdf_lst]
        pdf_merger = PdfFileMerger()
        for pdf in pdf_lst:
            pdf_merger.append(pdf)
        out_path = target_path + '/merge_pdf'
        if not os.path.exists(out_path):
            # 判断文件夹是否存在, 文件夹不存在就创建
            os.makedirs(out_path)
        pdf_merger.write(out_path + '/merge.pdf')


def get_opt_code():
    opt_code_dict = {
        '1': '1-PDF提取',
        '2': '2-PDF转图片',
        '3': '3-合并PDF'
    }
    opt_code = ''
    while opt_code == '':
        opt_code = input('请输入您的操作码: 1-PDF提取, 2-PDF转图片, 3-合并PDF')
        if opt_code_dict.get(opt_code) is None:
            print('您的输入[%s]有误，请重新输入' % opt_code)
            opt_code = ''
    print('您即将进行的操作是: %s, 请您按提示操作' % opt_code_dict[opt_code])
    return opt_code


def get_input_path(tips):
    ori_pdf_path = ''
    while ori_pdf_path == '':
        ori_pdf_path = input(tips)
        if os.path.exists(ori_pdf_path) is False:
            print("您的输入不存在, 请检查后重新输入")
            ori_pdf_path = ''
    return ori_pdf_path


def is_int(num):
    try:
        num = int(str(num))
        return isinstance(num, int)
    except:
        return False


def get_input_page():
    pdf_page_num = ''
    while pdf_page_num == '':
        pdf_page_num = input('请输入您要提取的PDF页码, 输入-1表示按页提取')
        if is_int(pdf_page_num) is False:
            print("您的输入不正确, 请输入整数(输入-1表示按页提取)")
            pdf_page_num = ''
    return int(pdf_page_num)


if __name__ == '__main__':
    opt_code = get_opt_code()
    pdf_editor = PdfEditor()
    if opt_code == '1':
        ori_pdf_path = get_input_path('请输入您的PDF的绝对路径: ')
        pdf_page_num = get_input_page()
        pdf_editor.get_page(ori_pdf_path, pdf_page_num)
        print('PDF提取完成，请在PDF所在目录的split_pdf目录下查看')
    elif opt_code == '2':
        ori_pdf_path = get_input_path('请输入您要转换的PDF的绝对路径: ')
        pdf_editor.pdf_to_image(ori_pdf_path)
        print('PDF转换图片完成，请在PDF所在目录的pdf_images目录下查看')
    else:
        print('请务必保证您的PDF文件名是有序的')
        ori_pdf_path = get_input_path('请输入您要合并的PDF的绝对路径: ')
        pdf_editor.merge_pdf(ori_pdf_path)
        print('PDF合并完成，请在PDF所在目录的merge_pdf目录下查看')
