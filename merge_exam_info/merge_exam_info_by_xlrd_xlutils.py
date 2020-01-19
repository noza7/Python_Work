from merge_exam_info.lib.func import get_filenames, get_students_exam_info_data, \
    from_computer_course_info_get_students_data
import xlrd
from xlutils.copy import copy

path = '计算机应用基础模板-.xlsx'
sheet_name = '计算机应用基础模板'

jsj_datas = from_computer_course_info_get_students_data(path=path)

# 中央开学生考试数据
# path = '签到表/中央开网考/data/'
path = 'xml/'
# 获取指定文件夹下所有xml文件名列表
filenames = get_filenames(path)[1:]
zyk_datas = get_students_exam_info_data(filenames, path)

print(f'计算机应用基础人数{len(jsj_datas)}')
print(f'网考人数{len(zyk_datas)}')

# 数据合并
for data in jsj_datas:
    zyk_datas.append(data)

print(f'合并后总人数{len(zyk_datas)}')

rb = xlrd.open_workbook('tzd.xls', formatting_info=True)
wb = copy(rb)
sheet = rb.sheet_by_name('sheet1')

# wb_w = xlwt.Workbook(encoding='utf-8')  # 创建workbook 对象
sheet_w = wb.get_sheet('sheet1')  # 创建工作表sheet

# 获取行数
rows = sheet.nrows
for a_i in range(1, rows):
    print(f'-------程序进行到第{a_i + 1}行-------')
    a_val = sheet.cell_value(a_i, 0)
    # print(a_val)
    try:
        stu_id = a_val[24:37]  # 学号
        print(stu_id)
        # 如果学号存在，从数据中遍历学生信息
        for i in range(len(zyk_datas) - 1, -1, -1):  # 倒叙遍历，为了可以删除找到的元素，缩短查找时间
            stu_info = zyk_datas[i]
            # 如果学号在信息表中找到
            if stu_id == stu_info[0]:
                # 遍历该学生所有考试科目，考试科目数量不会超过30
                for j in range(2, 30):
                    # 获取通知单试卷号
                    sjh = sheet.cell_value(rowx=a_i + j, colx=0)
                    # 先判断sjh是否为'考点名称：秦皇岛电大'，
                    if sjh == '考点名称：秦皇岛电大':
                        # print(sjh)
                        break
                    # 如果试卷号在列表中
                    elif str(sjh) in stu_info:
                        # 写入考场号
                        # wb['sheet1'][f'D{a_i + j}'].value = stu_info[1]
                        sheet_w.write(a_i + j, 3, stu_info[1])
                        # 写入座位号
                        sheet_w.write(a_i + j, 4, stu_info[2])
                        # 写入考试日期
                        sheet_w.write(a_i + j, 5, stu_info[4])
                        # 写入考试时间
                        sheet_w.write(a_i + j, 6, stu_info[5])
                        # print(stu_id)
                        print(f'{stu_id}试卷号{sjh}写入完毕！')
                        # 移除找到的元素，缩小查找范围
                        zyk_datas.pop(i)
        # print(len(zyk_datas))
    except Exception as e:
        print(e)
        continue
wb.save('test.xls')
