#!/bin/python3

import sys
import hashlib
import pathlib

import mammoth
from markdownify import markdownify

img_path:str = 'src'

def init_img_path():
    '''
    创建图片存储目录
    '''
    if pathlib.Path(img_path).exists(): return

    ## parents=True 表示如果目标目录的父级目录不存在的情况下自动创建, 
    ## 同 shell 中的 mkdir -p
    pathlib.Path(img_path).mkdir(parents = True)

def convert_img(img):
    '''
    对 word 文档中的图片进行处理, 将图片名直接转换成 md5 字符串
    '''
    img_suffix = img.content_type.split('/')[1]

    ## 这里只能用 with 语句, 如果改成 img_bytes = img.open() 会报错.
    with img.open() as img_bytes:
        img_byte_cnt = img_bytes.read()
        md5 = hashlib.md5()
        md5.update(img_byte_cnt)
        md5_str = md5.hexdigest()
        path_file = '{}/{}.{}'.format(img_path, md5_str, img_suffix)
        f = open(path_file, 'wb')
        f.write(img_byte_cnt)
        f.close()

    return {'src':path_file}

def doc2md(docx_path:str):
    init_img_path()

    ## 读取 Word 文件(二进制方式)
    docx_file = open(docx_path ,'rb')
    ## 获取无后缀的文件名称, 这里是为了防止文件名中包含多个点号, 所以没有直接使用[1]进行选择.
    docx_name = '.'.join(docx_path.split('.')[:-1])

    ## 转化 Word 文档为 HTML
    ## mammoth 有 convert_to_markdown() 方法
    result = mammoth.convert_to_html(docx_file, convert_image = mammoth.images.img_element(convert_img))
    docx_file.close()

    ## 获取 HTML 内容
    html_cnt = result.value
    ## 转化 HTML 为 Markdown
    md_cnt = markdownify(html_cnt, heading_style = 'ATX')
    md_file = open(docx_name + '.md', 'w', encoding='utf-8')
    md_file.write(md_cnt)
    md_file.close()

def doc2md_2(docx_path:str):
    '''
    使用 mammoth 库内置的 convert_to_markdown() 方法直接进行 docx -> markdown 的转换.
    不过转换出来的内容格式不如 markdownify 库简洁, 推荐使用 doc2md() 方法.
    '''
    init_img_path()
    ## 读取 Word 文件(二进制方式)
    docx_file = open(docx_path ,'rb')
    ## 获取无后缀的文件名称, 这里是为了防止文件名中包含多个点号, 所以没有直接使用[1]进行选择.
    docx_name = '.'.join(docx_path.split('.')[:-1])
    result = mammoth.convert_to_markdown(docx_file, convert_image = mammoth.images.img_element(convert_img))
    docx_file.close()

    md_file = open(docx_name + '.md', 'w', encoding='utf-8')
    md_file.write(result.value)

    md_file.close()

if __name__ == '__main__':
    ## python main.py d2m
    if len(sys.argv) < 3: 
        print("请指定操作类型")
        sys.exit(-1)
    if sys.argv[1] == 'd2m':
        doc2md(sys.argv[2])
    elif sys.argv[1] == 'm2d':
        pass
    else:
        pass
