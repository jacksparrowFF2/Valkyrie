#!/usr/bin/env python
# -*- encoding: utf-8 -*-
'''
@File    :   extract_raman.py
@Time    :   2019/09/19 16:59:18
@Author  :   SPH 
@Version :   1.0
@Contact :   s.ph@outlook.com
@License :   (C)Copyright 2018-2019, EXphysiclab
@Desc    :   None
'''

# here put the import lib
# 导入剪贴板组件
import win32clipboard as w
import win32con
# 导入命令行参数组件
import argparse
# 导入excel操作组件
import xlwings as xw
# 创建命令解释器
parser = argparse.ArgumentParser(description='This script is aims to extract Raman date from txt file')
# 创建命令行输入参数，输入参数为文件路径
# parser.add_argument("-i","--input", type=argparse.FileType(mode = 'r'), required = True, 
#                     help = 'the file need to extract data')
parser.add_argument("-i","--input", metavar = '', type=argparse.FileType(mode = 'r'),
                    help = 'the file need to extract data')
parser.add_argument("-e","--excel", metavar = '', type = str, help = 'the file need to open')
# 创建附属命令行参数，增加可选输出第二列的选项
group = parser.add_argument_group(description = 'Basic options')
group.add_argument('-c','--column', metavar = '', type = int, help = 'chose the column you want to extract')
# 创建互斥锁
# group = parser.add_mutually_exclusive_group()
group = parser.add_argument_group('exclusive options')
group.add_argument('-a','--all', action = 'store_true', 
                   help = 'this will extract all data to your clipboard')
group.add_argument('-s','--select', action = 'store_true', 
                   help = 'this will only extract the select column to your clipboard')
group.add_argument('-r','--write', action = 'store_true', 
                   help = 'this will add your raman data to your excel file last column')
group.add_argument('-ec','--condition', action = 'store_true', 
                   help = 'this will add the experiment condition store in your cilpboard to your excel last row')

args = parser.parse_args()

# 创建剪贴板读取函数
def getclip():
    w.OpenClipboard()
    copy_text = w.GetClipboardData(win32con.CF_UNICODETEXT)
    w.CloseClipboard()
    return copy_text

# 创建剪贴板写入函数
def writeclip(astring):
    w.OpenClipboard()
    w.EmptyClipboard()
    w.SetClipboardData(win32con.CF_UNICODETEXT, astring)
    w.CloseClipboard()
    
    
if __name__ == '__main__':
    # infile = args.input
    # filecontents = infile.read()
    
    if args.condition:
        # 获取剪贴板的内容
        excode_a = getclip()
        # 将剪贴板中的字符串从换行出切开，构建成列表
        excode_b = excode_a.split('\n')
        # excode_b = getclip().split('\n') #备选方式
        # 创建空列表
        excode_c = []
        # 从冒号处对列表中的每个元素的进行切割，构建成嵌套列表
        for i in excode_b:
            i = i.replace('\r', '')
            excode_c.append(i.split(':'))
        # 创建空字典
        excode = {}
        # 将嵌套列表转换成字典方便读取
        for i in range(len(excode_c)):
            excode[excode_c[i][0]] = excode_c[i][1]
        # # 调试输出-转化后的字典
        # print(excode)
        
        # 实验数据变量赋值
        date = excode["日期"]
        power = int(excode["初始输入功率(w)"])-int(excode["初始反馈功率(w)"])
        Ar = int(excode["Ar(sccm)"])
        H2 = int(excode["H2(sccm)"])
        CH4 = int(excode["CH4(sccm)"])
        pressure = int(excode["压强(pa)"])
        temp = int(excode["温度(℃)"])
        sub1 = excode["衬底1"]
        sub2 = excode["衬底2"]
        metaltype = excode["金属网"]
        note = excode["实验目的"]
        time = int(excode["持续时间(min)"])
        SR = int(excode["方阻(kΩ/□)"])
        
        # 创建实验条件数据列
        excondition = [note+"+"+sub1+"+"+sub2, metaltype, Ar, H2, CH4, time, power, pressure, temp, SR]
        # # 调试输出
        # print(excondition)
        # print(date)
        # 打开指定的EXCEL文件
        try:
            # 开始对 excel 文件进行操作
            inexcel = args.excel
            # 创建 App 进程
            app = xw.App(visible = False, add_book = False)
            # 链接工作表
            wb = app.books.open(inexcel)
            # 对指定工作表进行编辑
            sht = wb.sheets['Ratio Metadata']
            # 获取当前EXCEL表格的行数与列数
            info = sht.range('A1').expand('table')
            row = info.last_cell.row
            col = info.last_cell.column
            # 计算出要添加的一行位置
            rowl = row + 1
            # # 调试输出
            # print('原表格最后一行：'+str(row))
            # print('数据添加所在行：'+str(rowl))
            # 注入EXCEL公式
            sht.range('A'+str(rowl)).formula = '=B%s&CHAR(10)&AF%s'%(rowl,rowl)
            sht.range('K'+str(rowl)).formula = '=C%s/E%s'%(rowl,rowl)
            sht.range('L'+str(rowl)).formula = '=I%s/E%s'%(rowl,rowl)
            sht.range('M'+str(rowl)).formula = '=C%s/G%s'%(rowl,rowl)
            sht.range('N'+str(rowl)).formula = '=IF(88>J%s,45/(88-J%s),"bulk")'%(rowl,rowl)
            sht.range('Y'+str(rowl)).formula = '=Q%s*1.415'%(rowl)
            sht.range('Z'+str(rowl)).formula = '=R%s*1.01'%(rowl)
            sht.range('AA'+str(rowl)).formula = '=S%s*0.719'%(rowl)
            sht.range('AB'+str(rowl)).formula = '=Q%s&"/"&R%s&"/"&S%s'%(rowl,rowl,rowl)
            sht.range('AC'+str(rowl)).formula = '=Y%s/SUM(Y%s+Z%s+AA%s)'%(rowl,rowl,rowl,rowl)
            sht.range('AD'+str(rowl)).formula = '=Z%s/SUM(Y%s+Z%s+AA%s)'%(rowl,rowl,rowl,rowl)
            sht.range('AE'+str(rowl)).formula = '=AA%s/SUM(Y%s+Z%s+AA%s)'%(rowl,rowl,rowl,rowl)
            sht.range('AF'+str(rowl)).formula = '=P%s&"/"&Q%s&"/"&R%s&"/"&S%s&"/"&T%s&"/"&U%s&"/"&V%s&"/"&W%s'%(rowl,rowl,rowl,rowl,rowl,rowl,rowl,rowl)
            # 注入实验条件数据
            sht.range('O'+str(rowl), 'X'+str(rowl)).value = excondition
            sht.range('B'+str(rowl)).value = date
            # 对表格进行美化
                # 对第一行标题进行格式化
            sht.range('A1').expand('right').api.HorizontalAlignment = -4108
            sht.range('A1').expand('right').api.VerticalAlignment = -4108
                # 对数据列 A 进行格式化为水平居中对齐，字体加粗
            sht.range('A2').expand('down').api.HorizontalAlignment = -4152
            sht.range('A2').expand('down').api.VerticalAlignment = -4108
            sht.range('A2').expand('down').api.font.Bold = True
                # 对数据列 A 进行自动换行
            sht.range('A2').expand('down').api.WrapText = True
                # 对数据列 B 进行格式化为垂直居中对齐，水平靠右，加粗
            sht.range('B2').expand('down').api.HorizontalAlignment = -4152
            sht.range('B2').expand('down').api.VerticalAlignment = -4108
            sht.range('B2').expand('down').api.font.Bold = True
                # 对数据列 C:AE 进行格式化为水平垂直居中对齐
            sht.range('C2:AE'+str(rowl)).api.HorizontalAlignment = -4108
            sht.range('C2:AE'+str(rowl)).api.VerticalAlignment = -4108
                # 对数据列 B 进行格式化为垂直居中对齐，水平靠右，加粗
            sht.range('AF2').expand('down').api.HorizontalAlignment = -4152
            sht.range('AF2').expand('down').api.VerticalAlignment = -4108
                # 将数据列 AC:AE 数据显示格式化为百分比
            sht.range('AC2:AE2').expand('down').api.style = "Percent"
                # 将数据列 C:M,Y:AA 数据显示格式化为保留两位小数点
            sht.range('C2:J'+str(rowl)).api.NumberFormat = "##.00_)"
            sht.range('Y2:AA2').expand('down').api.NumberFormat = "##.00_)"
                # 将数据列 Q 数据显示格式化为 3 位
            sht.range('Q2').expand('down').api.NumberFormat = "000"
                # 将数据列 R:S 数据显示格式化为 2 位
            sht.range('R2:S2').expand('down').api.NumberFormat = "00"
                # 将数据列 T 数据显示格式化为 3 位
            sht.range('T2').expand('down').api.NumberFormat = "000"
            # 格式化完成提示
            print('EXCEL 格式化完成')
        finally:
            if wb:
                # 保存文件
                wb.save()
                # 关闭文件
                wb.close()
                # 结束进程
                app.kill
    elif args.write:
        infile = args.input
        filecontents = infile.read()
        print("you will write column 2 to your select excel file")
        # 转化至程序排序方式
        n = 1
        # 将字符串转换为列表，以换行符为切割处
        select_list = filecontents.split('\n')
        # 调试输出
        # print(select_list)
        # 构建格式化列表
        format_select_list = []
        for i in select_list:
            format_select_list.append(i.split('\t'))
        # 调试输出
        # print(format_select_list)
        # print(len(format_select_list))
        format_select_list.pop()
        # 调试输出
        # print(format_select_list)
        # print(len(format_select_list))
        # 构建输出列表
        out_select_list =[]
        for i in range(len(format_select_list)):
            # print(format_select_list[i][n])
            out_select_list.append(format_select_list[i][n])
            # out_select_list[i] = format_select_list[i][n]
        # # 调试输出
        # print(out_select_list)
        
        part1 = '=INDIRECT('
        part2 = '"'
        part3 = '\'Ratio Metadata\''
        part4 = '!$A'
        part5  = '"&COLUMN())'
        formula1 = part1 + part2 + part3 + part4 + part5
        print(formula1)
        
        out_select_list.insert(0,formula1)
        # # 调试输出
        # print(out_select_list)
        write_data = out_select_list
        # # 调试输出
        # print(write_data)
        try:
            # 开始对 excel 操作
            inexcel = args.excel
            # inexcel = 'r\'%s\'' %(args.excel)
            print('你输入的文件路径为：'+inexcel)
            # # 创建 app 进程
            # app = xw.App(visible = True, add_book = False)
            # 创建 app 进程
            app = xw.App(visible = False, add_book = False)
            # 链接工作表,填写要写入的EXCEL文件路径
            wb = app.books.open(inexcel)
            # 对指定工作表进行编辑
            sht = wb.sheets['Raman MetaData']
            # 获取当前EXCEL表格的行数与列数
            info = sht.range('A1').expand('table')
            # # 调试输出
            # print(info)
            row = info.last_cell.row
            col = info.last_cell.column
            # 计算出要添加的一行位置
            coll = col + 1
            str_coll = str(coll)
            print('数据添加所在列：'+str_coll)
            str_col = str(col)
            print('原表格最后一列：'+str_col)
            # 注入实验数据
            sht.range((1,coll),(row,coll)).options(transpose = True).value = write_data
            # 输出结果
            print('实验数据注入完成！')
            
            # 对新增列进行格式化
                # 对新增首行单元格进行格式化
            sht.range((1,coll)).api.HorizontalAlignment= -4152
            sht.range((1,coll)).api.VerticalAlignment= -4108
            sht.range((1,coll)).api.WrapText= True
                # 对新增数据单元格进行格式化
            sht.range((2,coll),(row,coll)).api.HorizontalAlignment= -4108
            sht.range((2,coll),(row,coll)).api.VerticalAlignment= -4108
                # 将新增单元格的列宽调整为 40
            sht.api.columns(coll).ColumnWidth = 40
        finally:
            if wb:
                #保存文件
                wb.save()
                # 关闭文件
                wb.close()
                # 结束进程
                app.kill
    elif args.all:
        infile = args.input
        filecontents = infile.read()
        print("this is all experiment data you get from test, you can find it in your clipborad")
        # # 调试输出
        # print(filecontents)
        writeclip(filecontents)
    elif args.select:
        infile = args.input
        filecontents = infile.read()
        print("you select column is %s" %(args.column))
        # 转化至程序排序方式
        n = args.column - 1
        # # 调试输出
        # print(n)
        # 将字符串转换为列表，以换行符为切割处
        select_list = filecontents.split('\n')
        # 调试输出
        # print(select_list)
        # 构建格式化列表
        format_select_list = []
        for i in select_list:
            format_select_list.append(i.split('\t'))
        # # 调试输出
        # print(format_select_list)
        # print(len(format_select_list))
        # 删除由 Raman 数据文件的最后一行空行
        format_select_list.pop()
        # # 调试输出
        # print(format_select_list)
        # print(len(format_select_list))
        
        # 构建输出列表
        out_select_list =[]
        for i in range(len(format_select_list)):
            # print(format_select_list[i][n])
            out_select_list.append(format_select_list[i][n])
            # out_select_list[i] = format_select_list[i][n]
        # print(out_select_list)
        # 构建输出字符串
        str_data = "\n".join(out_select_list)
        # # 调试输出
        # print(str_data)
        writeclip(str_data)
    else:
        # infile = args.input
        # filecontents = infile.read()
        # # 调试输出
        # print(filecontents)
        # writeclip(filecontents)
        print('请输入 -h 以查看使用说明')
        input("Press <enter>")
        

    # print(filecontents)
    
    # print(type(filecontents))
    # writeclip(filecontents)
    