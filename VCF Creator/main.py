import xlrd
import os
import sys
import pandas as pd

def create_info(name, number):
    text = 'BEGIN:VCARD' + '\n' \
        + 'VERSION:3.0' + '\n' \
        + 'PRODID:ez-vcard 0.11.2' + '\n'\
        + 'N;LANGUAGE=zh:;' + str(name) + '\n' \
        + 'FN:' + str(name) + '\n' \
        + 'TEL;TYPE=work:' + str(number) + '\n' \
        + 'END:VCARD' + '\n'
    return text

def process():
    #cur_file_path = os.path.dirname(os.path.realpath(__file__))
    abs_path = str(sys.argv[0])
    cur_file_path = str(os.path.abspath(abs_path))
    cur_file_path = cur_file_path[:cur_file_path.find('\main')]


    foler_list = os.listdir(cur_file_path)

    for stock_file in foler_list:
        file_name = stock_file[:stock_file.find('.')]
        tail = stock_file[stock_file.find('.'):]
        if(tail == '.xlsx'):

            path = os.path.join(cur_file_path, file_name + '.xlsx')
            df=pd.read_excel(path)#这个会直接默认读取到这个Excel的第一个表单
            
            result = ''
            for index, row in df.iterrows():
                name = df.at[index,'姓名']
                number = df.at[index, '手机号']
                text = create_info(name, number)
                result += text
            result = result[:-1]

            # 保存文件到本地
            result_path = os.path.join(cur_file_path,file_name + '.vcf')
            file = open(result_path,'w',encoding='utf-8')
            file.write(str(result))
            file.close()

if __name__ == '__main__':
    password = input('请输入爱的密码：')
    if('爱丹妮，爱天海' == password):
        process()
