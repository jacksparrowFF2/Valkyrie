# Valkyrie

 this program is aim to increase happiness

## 开始

### 安装 Quicker

> 可能是 Windows 最好的「Workflow」

软件官网：https://getquicker.net/

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
名称:值
```

例子：

```python
日期:20190607_01_00
实验目的:
实验过程:没有出现故障
初始输入功率(w):81
初始反馈功率(w):31
末端输入功率(w):86
末端反馈功率(w):36
Ar(sccm):150
H2(sccm):5
CH4(sccm):9
压强(pa):300
温度(℃):600
持续时间(min):120
衬底1:
衬底2:Quartz
金属网:MK0 镍网0.5_1.6
初步实验结果:
方阻(kΩ/□):1
```



## 程序

### 克隆仓库

将程序克隆到本地，使用编辑器对程序进行自定义编辑，并按照以下说明依次测试确保能够正确使用。

请注意，程序的正常运行依赖于模板，能够正确处理模板的信息时自动填写实验数据的第一步。

此外，为了降低因文件路径而引起的 BUG，请将程序放置于实验记录Excel同一文件夹内。

下面开始对三个程序进行简要的介绍和说明，希望你不会感到太多的疑惑。

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

import get
from get import app, info, row, rowl, sht, wb

global A, K, L, M, N, Y, Z, AA, AB, AC, AD

# 测试用变量 开始
# rowl = 10
# rowl = str(rowl)
# 测试用变量 结束

# A
# Excel原始公式
# =B84&CHAR(10)&AC84
# Excel公式定义
A = '=B'+rowl+'&CHAR(10)&AD'+rowl
# print(A)

# K ID/IG
# Excel原始公式
# =H84/D84
# Excel公式定义
K = '=C'+rowl+'/E'+rowl
# print(K)

# L IG'/IG
# Excel原始公式
# =B84/F84
# Excel公式定义
L = '=I'+rowl+'/E'+rowl
# print(L)

# M ID/ID'
# Excel原始公式
# =C84/G84
# Excel公式定义
M = '=C'+rowl+'/G'+rowl
# print(M)

# N 石墨烯层数
# Excel原始公式
# =IF(88>J84,45/(88-J84),"bulk")
# Excel公式定义
N = '=IF(88>J'+rowl+',45/(88-J'+rowl+'),"bulk")'
# print(J)

# Y 真实氩气
# Excel原始公式
# =Q84*1.415
# Excel公式定义
Y = '=Q'+rowl+'*1.415'
# print(Y)

# Z 氢气
# Excel原始公式
# =R84*1.01
# Excel公式定义
Z = '=R'+rowl+'*1.01'
# print(Z)

# AA
# Excel原始公式
# =S84*0.719
# Excel公式定义
AA = '=S'+rowl+'*0.719'
# print(AA)

# AB 气体流量比
# Excel原始公式
# =Q84&"/"&R84&"/"&S84
# Excel公式定义
AB = '=Q'+rowl+'&"/"&'+'R'+rowl+'&"/"&'+'S'+rowl
# print(AB)

# AC 真实气体流量
# Excel原始公式
# =Y84&"/"&Z84&"/"&AA84
# Excel公式定义
AC = '=Y'+rowl+'&"/"&'+'Z'+rowl+'&"/"&'+'AA'+rowl
# print(AC)

# AD TAG1
# Excel原始公式
# =P2&"/"&Q2&"/"&R2&"/"&S2&"/"&T2&"/"&U2&"/"&V2&"/"&W2
# Excel公式定义
AD = '=P'+rowl+'&"/"&'+'Q'+rowl+'&"/"&'+'R'+rowl+'&"/"&'+'S'+rowl+'&"/"&'+'T'+rowl+'&"/"&'+'U'+rowl+'&"/"&'+'V'+rowl+'&"/"&'+'W'+rowl
# print(AD)

# 检查公式是否正确，如果正确请注释

# print('A'+A)
# print('K'+K)
# print('L'+L)
# print('M'+M)
# print('N'+N)
# print('Y'+Y)
# print('Z'+Z)
# print('AA'+AA)
# print('AB'+AB)
# print('AC'+AC)
# print('AD'+AD)

# 如果程序调用正确，请注释以下所有内容
a = sht.range('A'+row).value
print(a)

#保存文件
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
import xlwings as xw
# 导入第一个程序
import get
# 导入变量
from excel_formula import AA, AB, AC, AD, A, K, L, M, N, Y, Z
from get import app, info, row, rowl, sht, wb

# 获取剪贴板内容
def getCopyText():
    wc.OpenClipboard()
    copy_text = wc.GetClipboardData(win32con.CF_UNICODETEXT)
    wc.CloseClipboard
    return copy_text


# 开始对剪贴板内容进行格式化，格式化为字典
excode_a = getCopyText()
excode_b = excode_a.split('\n')
excode_c = []
for i in excode_b:
    i = i.replace('\r', '')
    excode_c.append(i.split(':'))

excode = {}
for i in range(len(excode_c)):
    excode[excode_c[i][0]] = excode_c[i][1]
print(excode)

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
data = [note+"+"+sub1+"+"+sub2, metaltype, Ar, H2, CH4, time, power, pressure, temp, SR]
# print(data) # 验证数据列是否正确

# 注入实验条件数据
sht.range('O'+rowl, 'X'+rowl).value = data
sht.range('B'+rowl).value = date

# 注入Eecel公式
sht.range('A'+rowl).formula = A
sht.range('K'+rowl).formula = K
sht.range('L'+rowl).formula = L
sht.range('M'+rowl).formula = M
sht.range('N'+rowl).formula = N
sht.range('Y'+rowl).formula = Y
sht.range('Z'+rowl).formula = Z
sht.range('AA'+rowl).formula = AA
sht.range('AB'+rowl).formula = AB
sht.range('AC'+rowl).formula = AC
sht.range('AD'+rowl).formula = AD

# 保存文件
wb.save()
# 关闭文件
wb.close()
# 结束进程
app.kill

```

对上述程序配置完成后，请保存并进行测试，如果测试通过，复制要进行写入的实验条件，执行 `Write.py`，大约 5 秒后，写入完成。

## 解决方案

我对于这个项目希望它未来能做到全平台化(Win、Mac、Linux、IOS、Android)，能够集实验数据汇总、读取、云同步、案例对比和周汇报报告生成为一体。

目前因为个人能力有限，暂时只做到了实验数据的写入，以后会尝试写一个 UI 界面吧，不知道在研究生期间能不能完成，也许这个饼永远都不会完成。

我目前的解决方案：

+ 采用 ideaNote 进行实验记录的云同步
+ 使用 坚果云 同步所有的实验数据
+ Excel、Origin 进行数据分析

## 最后

希望这个小小的程序能够解决你的一部分痛点。