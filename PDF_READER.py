import PyPDF2
import argparse
import os
import pandas as pd
import sys
from Functions import adjust_excel_format


def extract_value(text, start_index):
    if text[start_index - 11] == " ":
        return 0.0
    value_str = text[start_index - 13:start_index - 3].strip()
    #strip()用于删除字符串前面和后面的所有空白字符
    return round(float(value_str), 3)

def collect_gas(pdf_reader, pdf_name):
    #为了让name在第一列进行的操作
    data = {key: 0.0 for key in Gas_Title}
    data_with_name = {"name": pdf_name}
    data_with_name.update(data)
    data = data_with_name
    page = pdf_reader.pages[1]
    text = page.extract_text()
    for idx, key in enumerate(Gas_Key):
        start_index = text.find(key)
        if start_index != -1:
            data[Gas_Title[idx]] = extract_value(text, start_index)
    return data

def parse_args():
    parser = argparse.ArgumentParser(description='Process PDF files')
    parser.add_argument('directory',help='Path to directory containing PDF files')
    return parser.parse_args()

def process_pdf_file(pdf_file_path):
    with open(pdf_file_path, 'rb') as pdf_file:
        pdf_reader = PyPDF2.PdfReader(pdf_file)
        i = -1
        for i in range(-1,-50,-1):
            if pdf_file_path[i]=="\\":
                break
        pdf_name = pdf_file_path[i + 1:-4]
        return collect_gas(pdf_reader, pdf_name)

def process_directory(directory_path):
    result = []
    for filename in os.listdir(directory_path):
        if filename.endswith('.pdf') and not filename.startswith('._'):
            pdf_file_path = os.path.join(directory_path, filename)
            print(pdf_file_path)
            try:
                data = process_pdf_file(pdf_file_path)
                result.append(data)
            except Exception as e:
                print(f"Error processing {pdf_file_path}: {e}")
    return result

def read_order_file(file_name):
    """读取文件内容，跳过以'#'开头的注释行"""
    if not os.path.isfile(file_name):
        raise FileNotFoundError("文件不存在,请检查TXT文件格式")
    with open(file_name, 'r', encoding='utf-8') as file:
        lines = file.readlines()
    #过滤掉开头为'#'的行数
    lines = [line for line in lines if not line.strip().startswith('#')]
    parts = lines[0].split(',')
    Gas_Key = [part for part in parts]
    parts = lines[1].split(',')
    Gas_Title = [part for part in parts]
    
    return Gas_Key,Gas_Title

    
if __name__ == '__main__':
    #只有该文件直接被执行的时候才会执行下面的模块，该文件被引用到别的py文件时name不是main
    if getattr(sys, 'frozen', False):
    # 如果是打包后的exe文件
        application_path = os.path.dirname(sys.executable)
    else:
    # 如果是普通Python脚本
        application_path = os.path.dirname(os.path.realpath(__file__))
    #current_directory = os.path.dirname(os.path.realpath(__file__))
    txt_path = os.path.join(application_path,"order.txt")
    Gas_Key,Gas_Title = read_order_file(txt_path)
    result = process_directory(application_path)
    #拼起来的list用dataframe可以直接变成表格形式
    print(application_path)
    df = pd.DataFrame(result)

    path_to_excel = os.path.join(application_path,"collect.xlsx")
    df.to_excel(path_to_excel, index = False)
    adjust_excel_format(path_to_excel)
    
    print(df)
    input("press any key")