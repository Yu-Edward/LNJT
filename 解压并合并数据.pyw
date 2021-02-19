import PySimpleGUI as sg,zipfile,os,shutil,xlrd,xlwt,sys
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

        if flag: [worksheet.write(0, k, label=v) for k, v in enumerate(sheet1.row_values(0))]
        flag = False
        # 插入其他文件内容
        for i in range(1, sheet1.nrows):
            [worksheet.write(current_row_number + i, k, label=v) for k, v in enumerate(sheet1.row_values(i))]
        current_row_number += sheet1.nrows - 1
    else:
        workbook.save(path)
    return 0

flag = True
time_num = 0
while True:
    filename_dir = sg.popup_get_file('请选择文件：')

    if filename_dir is None :
        # 输入框选择的是'Cancel'，或者是'X'
        sg.popup('SB !!!')
        sys.exit()
    elif not filename_dir:
        # 没有选择文件，直接点击'OK'
        time_num += 1
        if time_num == 3:
            sg.popup('SB !!!')
            sys.exit()
        continue
    else:
        break

output_path = '/'.join(filename_dir.split("/")[:-1]) + '/'
filename = filename_dir.split("/")[-1]


makeUnzipOfFolder(output_path,filename)
copyDocument(filename_dir.replace('.zip','') + '/',filename.replace('.zip','.xls'),filename_dir.replace('.zip',''))

sg.popup('Finish !!!',keep_on_top = True)