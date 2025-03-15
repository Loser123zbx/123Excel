
from _decimal import Decimal  
import os
import random
from pdf2docx import Converter
import xml.etree.ElementTree as et
import zipfile 
import os
import json
import decimal
import pandas as pd

def read_config():
    with open("123XLS.json","r",encoding = "utf-8") as f:
        return  json.load(f)

config = read_config()
#round_way
RW = config['using_round_way']
#decimal_places_num 
DPN = config['decimal_places_num']

CORE = config['core']

def read_excel(excel_file,output_dir = 'temp_xml', sheet = ''):

    if CORE == '123Excel':
        sheet = sheet + ".xml"

        is_xlsx = excel_file.split('.')[-1] == 'xlsx'
        is_xls = excel_file.split('.')[-1] == 'xls'

        if is_xlsx:
            if not os.path.exists(output_dir):
                os.makedirs(output_dir)

            with zipfile.ZipFile(excel_file, 'r') as zip_ref:

                    file_names = zip_ref.namelist()
                    sheet_files = []
                    for i in file_names:
                        if i.split("/")[0] == 'xl' and i.split("/")[1] == "worksheets":
                            sheet_files.append(i)

                    for xml_file in sheet_files:
                        xml_content = zip_ref.read(xml_file)

                        output_path = os.path.join(output_dir, os.path.basename(xml_file))

                        # 写入到本地文件
                        with open(output_path, 'wb') as output_file:
                            output_file.write(xml_content)

            tree = et.parse(os.path.join(output_dir, sheet))
            root = tree.getroot()
            rows_data = []

            # 查找所有的<row>元素
            rows = root.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}row')
            for row in rows:
                # 初始化一个列表来存储当前行的数据
                row_data = []

                # 查找当前行中的所有<c>元素
                cells = row.findall('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}c')
                for cell in cells:
                    # 获取<v>标签的值
                    v_value = cell.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}v')
                    if v_value is not None:
                        row_data.append(v_value.text)
                    else:
                        is_value = cell.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}is/{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t')
                        if is_value is not None:
                            row_data.append(is_value.text)
                        else:
                            row_data.append(None)
                # 将当前行的数据添加到总列表中
                rows_data.append(row_data)
            os.remove(os.path.join(output_dir,sheet))
            return rows_data
        if is_xls:
            df = pd.read_excel(excel_file, sheet_name=sheet)

            # 将 DataFrame 转换为二维列表
            data = df.values.tolist()

            return data
    elif CORE == 'pandas':
        df = pd.read_excel(excel_file, sheet_name=sheet, header=None)
        # 将 DataFrame 转换为二维列表
        data = df.values.tolist()
        return data
def find_student_score(value,student_arrange,subject_row = 1):
    students = []
    for i in value[(student_arrange-1):]:
        students.append(i[student_arrange-1])
    return students

def get_subjects(value,subject_row = 1):
    return value[subject_row-1]

def zidingyi(value,sub,student_row,min,max,subject_row = 1):
    sub_index = value[subject_row-1].index(sub)
    students=[]
    scores=[]
    for i in value:
        try:
            if int(i[sub_index]) >= int(min) and int(i[sub_index]) <= int(max):
                students.append(i[student_row-1])
                scores.append(i[sub_index])
        except:
            pass
    num = len(students)
    output_message = []
    output_message.append(f"{sub}:{min}分>>{max}分的学生有{num}:")
    for i in range(num):
        output_message.append(f"{students[i]}的分数是{scores[i]}")
    return output_message

def get_avg(value,sub,subject_row = 1):
    '''
    ROUND_UP: 远离零方向舍入，即总是增加绝对值。
    //向上取整

    示例：Decimal('1.3').quantize(Decimal('1'), rounding=ROUND_UP) 结果为 2
    示例：Decimal('-1.3').quantize(Decimal('1'), rounding=ROUND_UP) 结果为 -2

    ROUND_DOWN: 靠近零方向舍入，即总是减少绝对值。
    //向下取整

        示例：Decimal('1.7').quantize(Decimal('1'), rounding=ROUND_DOWN) 结果为 1
        示例：Decimal('-1.7').quantize(Decimal('1'), rounding=ROUND_DOWN) 结果为 -1
    
    ROUND_CEILING: 向正无穷方向舍入，只对负数有效。[不必要]
    //负数向上取整

        示例：Decimal('-1.1').quantize(Decimal('1'), rounding=ROUND_CEILING) 结果为 -1
    
    ROUND_FLOOR: 向负无穷方向舍入，只对正数有效。
    //正数向下取整
    
        示例：Decimal('1.1').quantize(Decimal('1'), rounding=ROUND_FLOOR) 结果为 1
         
    ROUND_HALF_UP: 如果舍弃部分 >= 0.5，则向上舍入；否则向下舍入。
    //四舍五入

        示例：Decimal('1.5').quantize(Decimal('1'), rounding=ROUND_HALF_UP) 结果为 2
        示例：Decimal('2.4').quantize(Decimal('1'), rounding=ROUND_HALF_UP) 结果为 2
    
    ROUND_HALF_DOWN: 如果舍弃部分 > 0.5，则向上舍入；否则向下舍入。
    //五舍六入

        示例：Decimal('1.5').quantize(Decimal('1'), rounding=ROUND_HALF_DOWN) 结果为 2
        示例：Decimal('2.5').quantize(Decimal('1'), rounding=ROUND_HALF_DOWN) 结果为 2
    
    ROUND_HALF_EVEN: 银行家舍入法，如果舍弃部分 = 0.5，则选择最接近的偶数。
    //银行家舍入

    示例：Decimal('2.5').quantize(Decimal('1'), rounding=ROUND_HALF_EVEN) 结果为 2
    示例：Decimal('3.5').quantize(Decimal('1'), rounding=ROUND_HALF_EVEN) 结果为 4
    '''
    subjects = get_subjects(value,subject_row)
    sub_index = subjects.index(sub)
    
    sumA = 0
    for i in value: 
        try:
            add_value = decimal.Decimal(str(i[sub_index]))
            sumA = decimal.Decimal(sumA + add_value)
        except:
            sumA += 0
    sumA = decimal.Decimal(str(sumA))
    count = decimal.Decimal(str(len(value)))
    average = decimal.Decimal(sumA / count)
    dpn = str( 1 / (10 ** int(DPN)) )
    output_message = f"{sub}平均分 {average.quantize(decimal.Decimal(dpn), rounding = RW)}"
    return output_message



def get_max(value,sub,student_row,subject_row = 1):
    sub_index = value[subject_row-1].index(sub)
    max = 0
    max_student = []
    for i in value[1:]:
        try:
            if float(i[sub_index]) > max:
                max = float(i[sub_index])
        except:
            raise ValueError(f"line {i} isn't a float!")
    for i in value:
        try:
            if float(i[sub_index]) == max:
                max_student.append(i[student_row-1])
        except:
            pass
    output_message = []
    output_message.append(f"{sub}最高分{max}:")
    for i in max_student:
        output_message.append(i)
    return output_message
def get_min(value,sub,student_row,subject_row = 1):
    sub_index = value[subject_row-1].index(sub)
    min = 0
    min_student = []
    for i in value:
        try:
            if float(i[sub_index]) < min:
                min = float(i[sub_index])
        except:
            pass
    for i in value:
        try:
            if float(i[sub_index]) == min:
                min_student.append(i[student_row-1])
        except:
            pass
    output_message = []
    output_message.append(f"{sub}最低分{min}:")
    for i in min_student:
        output_message.append(i)
    return output_message

def get_student_score(value,student_row,name,subject_row = 1):
    subjects = get_subjects(value,subject_row-1)
    for i in value:
        if i[student_row-1] == name:
            score = i
    output_message = []
    output_message.append(f"{name}的分数是:")
    for i in range(len(subjects)):
        output_message.append(f"{subjects[i]}:{score[i]}")
    return output_message
