import PySimpleGUI as sg,zipfile,os,shutil,xlrd,xlwt
from pathlib import Path



def makeUnzipOfFolder(filedir:str,filename:str):
    '''
    解压缩文件，并解压缩到当前目录下生成以压缩文件名命名的新目录下
    :param filedir: 压缩文件所在的目录路径
    :param filename: 压缩文件名
    :return:
    '''
    # 创建加压缩后的目录；如果目录存在，删除后重新创建
    path = filedir + filename.replace('.zip','')
    if os.path.exists(path):
        shutil.rmtree(path)
    else:
        os.makedirs(path)
    file_name = filedir + filename
    zipf = zipfile.ZipFile(file_name)
    zipf.extractall(path=path)
    zipf.close()
    return 0

def removeDir(file_dir):
    if os.path.exists(file_dir):
        shutil.rmtree(file_dir)
    else:
        return 1
    return 0

def copyDocument(file_dir, filename, source_dir):
    '''
    指定目录下的xls文件中的内容合并到指定的xls文件中
    :param file_dir: 合并后文件目录
    :param filename: 合并后文件名
    :param source_dir: 源文件的目录路径
    :return: 0 成功
    '''
    path = file_dir + filename

    p = Path(source_dir)
    xls_file_dir = [i for i in p.iterdir()]

    # 创建xls
    workbook = xlwt.Workbook()
    # 创建sheet
    worksheet = workbook.add_sheet('第1页', cell_overwrite_ok=True)

    # 设定初始行数为0，后将文件的行数累计相加
    current_row_number = 0
    flag = True
    for xls_file_name_dir in xls_file_dir:
        # 依次打开源文件
        workbook1 = xlrd.open_workbook(xls_file_name_dir)
        sheet1 = workbook1.sheet_by_name(workbook1.sheet_names()[0])

        # 第一行内容，并插入到xls
        # firstRow = ['车牌号码', '车牌颜色', '车辆属地管辖机构名称', '所属行业', '所属地区', '车辆类型', '业户名称', '平台名称', '入网状态', '年审截止日期', '第一次入网时间',
        #             '最后上线日期']
        if flag: [worksheet.write(0, k, label=v) for k, v in enumerate(sheet1.row_values(0))]
        flag = False
        # 插入其他文件内容
        for i in range(1, sheet1.nrows):
            [worksheet.write(current_row_number + i, k, label=v) for k, v in enumerate(sheet1.row_values(i))]
        current_row_number += sheet1.nrows - 1
    else:
        workbook.save(path)
    return 0


file_list = ('未入网、未上线数据','轨迹漂移率、轨迹完成率数据','数据合格率低于99.99%的车辆数据','报警数据')
file_list_detail = {}
for i in file_list:
    # zip文件及其路径,例 d://Users//Administrator//Desktop//123//1610334284753889.zip
    filename_dir = sg.popup_get_file('请选择 \'{}\' 文件'.format(i))
    # zip文件所在目录,例 d://Users//Administrator//Desktop//123//
    output_path = '/'.join(filename_dir.split("/")[:-1]) + '/'
    # zip文件名，例 1610334284753889.zip
    filename = filename_dir.split("/")[-1]
    file_list_detail[i] = [filename_dir,output_path,filename]

# print(filename_dir)
# print(output_path)
# print(filename)

[makeUnzipOfFolder(file_list_detail[i][1],file_list_detail[i][2]) for i in file_list]


# 解压缩后的目录路径
unzip_dir = output_path + filename.replace('.zip','')
# print(unzip_dir)

# p = Path(unzip_dir)
# xls_file = [ i for i in p.iterdir()]
# print(xls_file[0])

[copyDocument('D://专项整治//数据表//', '总表{}.xls'.format(k + 1), file_list_detail[v][1] + file_list_detail[v][2].replace('.zip','')) for k,v in enumerate(file_list)]
[removeDir(file_list_detail[v][1] + file_list_detail[v][2].replace('.zip','')) for k,v in enumerate(file_list)]


# 行政审批局车辆.xls
path = 'D://专项整治//数据表//总表1.xls'
workbook = xlrd.open_workbook(path)
sheet = workbook.sheet_by_name(workbook.sheet_names()[0])
nrows = sheet.nrows

workbook1 = xlwt.Workbook()
worksheet = workbook1.add_sheet('第1页', cell_overwrite_ok=True)

# 写入第一行
[worksheet.write(0, k, label=v) for k, v in enumerate(sheet.row_values(0))]

# 行政审批局车辆写入表中并保存
rows_number = 1
for i in range(1,nrows):
    if sheet.row_values(i)[2] == '行政审批局' and '营口' in sheet.row_values(i)[4]:
        [worksheet.write(rows_number, k, label=v) for k, v in enumerate(sheet.row_values(i))]
        rows_number += 1
else:
    workbook1.save('D://专项整治//数据表//行政审批局车辆.xls')


sg.popup("Finish!")

import zxzz