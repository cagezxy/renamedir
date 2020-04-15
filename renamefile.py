# coding=utf-8
import os
import re
from operateExcel import ExcelUtil


def rename_file(filepath, excelpath, sheetname):
    data = ExcelUtil(excelpath, sheetname).dict_data()
    # print(data)
    print("excel 编号长度： {}".format(len(data)))
    print("切换到目录： {}".format(filepath))
    os.chdir(filepath)
    print("开始重命名\n")
    if not os.path.exists(filepath) or not os.path.isdir(filepath):
        print("文件目录不存在 或者 文件所给的路径不是文件夹")
        os._exit(1)
    filenames = os.listdir(filepath)
    i = 0
    for fn in filenames:
        if not os.path.isdir(fn):
            print("{:15s}不是文件夹，跳过...".format(fn))
            continue
        # # 文件名前缀
        # file_name_tmp = os.path.splitext(fn)[0]
        # # 文件名后缀
        # file_ext_name = os.path.splitext(fn)[1]
        file_name_tmp = fn
        # 找到中文名,取找到的第一个
        fn_tmp_name = re.compile(r'[\u4e00-\u9fa5]+').findall(file_name_tmp)
        if len(fn_tmp_name) == 0:
            print("{:15s} 目录没有找到中文名".format(fn))
            continue
        name_key = fn_tmp_name[0]
        if name_key in data:
            newname = str(data[name_key]) + "-" + name_key
            print("{:15s}重命名为---->{}".format(fn, newname))
            if os.path.exists(newname):
                print("{:15s}已经存在，请检查，跳过...".format(newname))
                continue
            # rename 中path已经在当前路径下， 则参数直接使用 文件名就行
            os.rename(fn, newname)
            i += 1
        else:
            print("{:15s} 目录没有找到对应中文名".format(fn))
            continue
    print("\n重命名结束，成功 {} 个，请检查。".format(i))


if __name__ == '__main__':
    print("""使用说明：
请输入需要重命名文件夹 所在目录的上一级目录:
        格式为： D:\\示例\\杭州分公司1-5\\
请输入excel所在完整路径：
        格式为： D:\\示例\\浙江大区：常温事业部管理人员廉洁档案、廉洁承诺书汇总表（20年）(2).xlsx
请输入 excel sheet页名字，不输入默认Sheet1：
        格式为： Sheet1
    """)
    file_path = input("请输入需要重命名文件夹 所在目录的上一级目录： ")
    excel_path = input("请输入excel所在完整路径: ")
    sheet_name = input("请输入 excel sheet页名字: ")
    sheet_name = sheet_name.strip()
    if len(file_path) == 0 or file_path is None:
        # 当然了， 如果 有几个路径需要修改，放入[] 中， 然后循环就行
        # file_path = "D:\\needtorename\\示例\\杭州分公司1-5\\"
        file_path = "D:\\办公企划专员\\2020\\浙江大区廉洁档案4.13\\廉洁档案\\档案\\浙江大区1-354\\"
    if len(excel_path) == 0 or excel_path is None:
        excel_path = "D:\\needtorename\\示例\\浙江大区：常温事业部管理人员廉洁档案、廉洁承诺书汇总表（20年）(2).xlsx"
    if len(sheet_name) == 0 or sheet_name is None:
        sheet_name = "Sheet1"
    try:
        rename_file(file_path, excel_path, sheet_name)
    except Exception as e:
        print("重命名失败，报错{}".format(e))
        input("任意按键结束： ")
        os._exit(1)
    input("任意按键结束： ")
    os._exit(0)