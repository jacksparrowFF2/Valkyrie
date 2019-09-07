# Valkyrie

 this program is aim to increase happiness

## 开始

## 安装 Quicker

> 可能是 Windows 最好的「Workflow」、

软件官网：https://getquicker.net/

模板处理动作：



复制到软件并粘贴：



### 安装 Scoop

> 可能是 Windows 上体验最好的「包管理器」

**如果你的电脑上已经安装了 Python3 请跳过这一节。**

Scoop 的安装需要满足以下基础条件：

+ Windows 版本不低于 Windows 7
+ Windows 中的 PowerShell 版本不低于 PowerShell 3
+ 你能 **正常、快速** 的访问 GitHub 并下载上面的资源
+ 你的 Windows 用户名为英文（Windows 用户环境变量中路径值不支持中文字符）

按下键盘上的`win + x`键，在弹出的菜单中选择 Windows PowerShell:

```
# 在 PowerShell 中输入下面内容，保证允许本地脚本的执行
set-executionpolicy remotesigned -scope currentuser
# 执行以下命令安装 Scoop
iex (new-object net.webclient).downloadstring('https://get.scoop.sh')
# 等待脚本执行完成并验证是否安装成功
scoop help
```

如果输出以下结果则代表安装成功：

```
Usage: scoop <command> [<args>]

Some useful commands are:

alias       Manage scoop aliases
bucket      Manage Scoop buckets
cache       Show or clear the download cache
checkup     Check for potential problems
cleanup     Cleanup apps by removing old versions
config      Get or set configuration values
create      Create a custom app manifest
depends     List dependencies for an app
export      Exports (an importable) list of installed apps
help        Show help for a command
hold        Hold an app to disable updates
home        Opens the app homepage
info        Display information about an app
install     Install apps
list        List installed apps
prefix      Returns the path to the specified app
reset       Reset an app to resolve conflicts
search      Search available apps
status      Show status and check for new app versions
unhold      Unhold an app to enable updates
uninstall   Uninstall an app
update      Update apps, or Scoop itself
virustotal  Look for app's hash on virustotal.com
which       Locate a shim/executable (similar to 'which' on Linux)


Type 'scoop help <command>' to get help for a specific command.
```

### 安装 Python3

在确认 Scoop 安装完成后，打开 PowerShell：

```
Scoop install python
```

等待安装完成并尝试在 PowerShell 中输入:

```
python
```

如果输出以下信息则代表安装正确：

```
Python 3.7.4 (tags/v3.7.4:e09359112e, Jul  8 2019, 20:34:20) [MSC v.1916 64 bit (AMD64)] on win32
Type "help", "copyright", "credits" or "license" for more information.
```

###  安装编辑器

```
# 安装 VSCode
scoop install vscode
# 安装 Sublime
scoop install sublime-text
# 安装 NotePad
scoop install notepadplusplus-np
```

以上三个软件任选其一，但是我推荐首选用 VSCode，这三个软件的主要用途就是编辑代码，这里不做过多展开。

### 安装依赖

```
# Python 中用于对 EXCEL 进行操作的组件
pip3 install xlwings
# Python 中用于访问 Window 剪贴板的组件
pip3 install pywin32
```

至此，程序所需运行的环境以及依赖都已准备完毕。

## 模板

模板的创建时为了程序能够正常的读取并写入到数据库中，在此给出模板要遵循的格式：

```python
{
	中文/英文/中文+英文/纯数字/符号:中文/英文/中文+英文/纯数字/符号
}
```

例子：

```python
{
	日期:20190907_01_01
	date:20190907_01_01
	实验目的:改变
	初始输入功率(w):81
	Ar(sccm):100
	衬底2:Quartz
	金属网:MK1 铜网1_1
	初步实验结果:ssAA啊啊啊
	方阻(kΩ/□):1
}
```

为了确保能够将模板中的内容处理为 Python 能够识别的格式，可以先用 正则表达式 对其进行匹配并测试。

```python
{
	"日期":"20190907_01_01",
	"date":"20190907_01_01",
	"实验目的":"改变",
	"初始输入功率(w)":"81",
	"Ar(sccm)":"100",
	"衬底2":"Quartz",
	"金属网":"MK1 铜网1_1",
	"初步实验结果":"ssAA啊啊啊",
	"方阻(kΩ/□)":"1",
}
```



## 程序

请注意，程序的正常运行依赖于模板，能够正确处理模板的信息时自动填写实验数据的第一步。

### get.py

这一段程序主要是为了获取 EXCEL 表格中已有的数据阵列并计算出应填写数据的所在位置，详细含义请参见程序注释。

```python
# 引入模块
import xlwings as xw
# 定义全局变量便于其他程序调用
global app, wb, sht, info, row, rowl
# 开始对EXCEL进行编辑

# 创建app进程
app = xw.App(visible=False, add_book=False)
# 链接工作表
wb = app.books.open('填写要修改的 EXCEl 的文件地址')#填写要写入的EXCEL文件路径
# 对指定工作表进行编辑
sht = wb.sheets['填写要修改的工作表，例如 Sheet1']
# 获取当前EXCEL表格的行数与列数
info = sht.range('A1').expand('table')
print(info)
# 计算出当前表格最后一行
row = info.last_cell.row
# 计算出当前表格最后一列
col = info.last_cell.column
# 计算并出要添加的一行位置
rowl = str(row + 1)
print('数据添加所在行：'+rowl)
# 计算并输出原表格的最后一行
row = str(row)
print('原表格最后一行：'+row)
```

### Excel_Formula.py

这一段程序是为了向 Excel 单元格中注入公式，因为不会使用 VBA 也没有找到在 Python 中用 VBA 计算 Excel 的方法，所以只能采取这种笨办法，如果你有更好的方案，欢迎您提出建议。

```python
#!/usr/bin/env python
# -*- encoding: utf-8 -*-

from get import app, wb, sht, info, row, rowl
import get
# 定义全局变量，建议用 EXCEL 的列编号定义变量名称
global A, J, K, L, X, Y, Z, AA, AB, AC

# 测试用变量 开始
# rowl = 10
# rowl = str(rowl)
# 测试用变量 结束

# A
# Excel原始公式
# =AD84&CHAR(10)&AC84
# Excel公式定义
A = '=AD'+rowl+'&CHAR(10)&AC'+rowl
# print(A)

# J
# Excel原始公式
# =B84/D84
# Excel公式定义
J = '=B'+rowl+'/D'+rowl
# print(J)

# K
# Excel原始公式
# =H84/D84
# Excel公式定义
K = '=H'+rowl+'/D'+rowl
# print(K)

# L
# Excel原始公式
# =B84/F84
# Excel公式定义
L = '=B'+rowl+'/F'+rowl
# print(L)

# M
# Excel原始公式
# =IF(88>I84,45/(88-I84),"bulk")
# Excel公式定义
M = '=IF(88>I'+rowl+',45/(88-I'+rowl+'),"bulk")'
# print(M)

# X
# Excel原始公式
# =P84*1.415
# Excel公式定义
X = '=P'+rowl+'*1.415'
# print(X)

# Y
# Excel原始公式
# =Q84*1.01
# Excel公式定义
Y = '=Q'+rowl+'*1.01'
# print(Y)

# Z
# Excel原始公式
# =R84*0.719
# Excel公式定义
Z = '=R'+rowl+'*0.719'
# print(Z)

# AA
# Excel原始公式
# =P84&"/"&Q84&"/"&R84
# Excel公式定义
AA = '=P'+rowl+'&"/"&'+'Q'+rowl+'&"/"&'+'R'+rowl
# print(AA)

# AB
# Excel原始公式
# =X84&"/"&Y84&"/"&Z84
# Excel公式定义
AB = '=X'+rowl+'&"/"&'+'Y'+rowl+'&"/"&'+'Z'+rowl
# print(AB)

# AC
# Excel原始公式
# =O84&"/"&P84&"/"&Q84&"/"&R84&"/"&S84&"/"&T84&"/"&U84&"/"&V84
# Excel公式定义
AC = '=O'+rowl+'&"/"&'+'P'+rowl+'&"/"&'+'Q'+rowl+'&"/"&'+'R'+rowl + \
    '&"/"&'+'S'+rowl+'&"/"&'+'T'+rowl+'&"/"&'+'U'+rowl+'&"/"&'+'V'+rowl
# print(AC)

#检查公式是否正确，如果输出正确请注释
print('A'+A)
print('J'+J)
print('K'+K)
print('L'+L)
print('X'+X)
print('Y'+Y)
print('Z'+Z)
print('AA'+AA)
print('AB'+AB)
print('AC'+AC) 

# 测试第一个程序调用是否正常，如果调用正常请注释
a = sht.range('A'+row).value
print(a)

# 如果这个程序执行正常，请将下面这段程序注释

# 保存文件
wb.save()
# 关闭文件
wb.close()
# 结束进程
app.kill


```

### Write.py

```python
#!/usr/bin/env python
# -*- encoding: utf-8 -*-

# 导入剪贴板相关模块
import win32clipboard as wc
import win32con
# 导入 ast,用于将字符串类型转变为字典类型
import ast
# 导入 Excel 操作相关模块
import xlwings as xw
# 导入第一个程序
import get
# 导入变量
from excel_formula import A, J, K, L, X, Y, Z, AA, AB, AC
from get import app, wb, sht, info, row, rowl

# 获取剪贴板内容
def getCopyText():
    wc.OpenClipboard()
    copy_text = wc.GetClipboardData(win32con.CF_UNICODETEXT)
    wc.CloseClipboard
    return copy_text


# 将字符串类型转变为字典类型
excode = ast.literal_eval(getCopyText())

# 输出变量类型，确保为字典类型
print(type(excode))

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
data = [note, metaltype, Ar, H2, CH4, time, power, pressure, temp, SR]
# 验证数据列是否正确，如果正确请注释下面一行程序
print(data) 

# 注入实验条件数据
sht.range('N'+rowl, 'W'+rowl).value = data
sht.range('AD'+rowl).value = date

# 注入Eecel公式
sht.range('A'+rowl).formula = A
sht.range('J'+rowl).formula = J
sht.range('K'+rowl).formula = K
sht.range('L'+rowl).formula = L
sht.range('X'+rowl).formula = X
sht.range('Y'+rowl).formula = Y
sht.range('Z'+rowl).formula = Z
sht.range('AA'+rowl).formula = AA
sht.range('AB'+rowl).formula = AB
sht.range('AC'+rowl).formula = AC

# 保存文件
wb.save()
# 关闭文件
wb.close()
# 结束进程
app.kill

```

