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

**可是，如果我的电脑不支持 Scoop 怎们办？没有关系，对应的软件都可以在官网找到，只不过通过 Scoop 进行安装更为方便与快捷。**

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

**请注意，如果你的实验记录里，标注了某个名称但是没有相对应的值，程序在接下来处理可能会出现错误。**

### 安装编辑器

```
# 安装 VSCode
scoop install vscode
# 安装 Sublime
scoop install sublime-text
# 安装 NotePad
scoop install notepadplusplus-np
```

以上三个软件任选其一，但是我推荐首选用 VSCode，这三个软件的主要用途就是编辑代码，这里不做过多展开。

## 程序

### 克隆仓库

将程序克隆到本地，使用编辑器对程序进行自定义编辑，并对修改的内容依次测试确保能够正确使用。

~~此外，为了降低因文件路径而引起的 BUG，请将程序放置于实验记录Excel同一文件夹内。~~为了解决和避免这一情况的出现（如果变更文件路径则需要频繁的打开程序并修改，还有可能会引起不必要的 Bug），所以从新编写了程序，使其更为方便、与快捷。

现在，你不需要知道每一行代码的含义（事实上我也不知道，我估摸着能用），你只需要结合 Quicker 这个用于提升工作效率神器，将鼠标放在合适的位置并进行正确的点击即可。当然你仍然可以对程序进行编辑，以符合你自身的情况。

我将从新编写好的程序放在了 `data_extract`这个文件夹内，`idea`文件夹是我用于测试以及模板，测试数据的存放位置，`history`是存放之前用于解决问题的程序，用于存档和备忘。目前在`data_extract`文件夹内，只有两个文件：

+ extract_JV
+ extract_raman

未来会随着自身实验的推进和与师兄师姐的交流不断添加、优化新的程序以处理不同测试设备导出的数据和满足文件处理的要求。

### 运行程序

#### 通过命令行

> 虽然在这里介绍了如何通过命令行运行程序以得到自己期待的结果，但最终，我想大多数人都会采用另一种方式——通过 Quicker 进行调用。

将

下面开始对三个程序进行简要的介绍和说明，希望你不会感到太多的疑惑。

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