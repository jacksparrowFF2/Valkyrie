Valkyrie

 这一项目旨在提升处理数据时幸福指数以及节省时间（并没有）。

# 开始

### 项目地址

Github： https://github.com/jacksparrowFF2/Valkyrie 

Gitee： https://gitee.com/nekosan/Valkyrie 

## 待做

- [ ] 图形化界面
- [ ] 数据处理需求征集

## 安装 Quicker（必选）

> 可能是 Windows 最好的「Workflow」

软件官网：https://getquicker.net/

## 安装 Scoop（可选）

> 可能是 Windows 上体验最好的「包管理器」

**如果你的电脑上已经安装了 Python3 ，请跳过这一节。**

**如果你对于命令行这一概念不熟悉，请跳过这一节。**

**如果你的电脑所在的网络环境并不能正常的或者流畅的访问国外网站，请跳过这一节。**

首先。Scoop 的安装需要满足以下基础条件：

+ Windows 版本不低于 Windows 7
+ Windows 中的 PowerShell 版本不低于 PowerShell 3
+ 你能 **正常、快速** 的访问 GitHub 并下载上面的资源
+ 你的 Windows 用户名为英文（Windows 用户环境变量中路径值不支持中文字符）

**可是，如果我的电脑不支持 Scoop 怎们办？没有关系，对应的软件都可以在官网找到，只不过通过 Scoop 进行安装更为方便与快捷。**

按下键盘上的`win + x`键，在弹出的菜单中选择 Windows PowerShell:

```powershell
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
## 安装编辑器（必选）

如果你能顺利的安装 Scoop 以及访问国外网站，为了节约你的时间，请通过 Scoop 复制相应的命令以安装下方任意一个编辑器：

- 安装 VSCodeVSCode

```powershell
scoop install vscode
```

- 安装 Sublime

```
scoop install sublime-text
```

- NotePad++

```powershell
scoop install notepadplusplus-np
```

如果你不能顺利的安装 Scoop 以及访问国外网站，请移步至相应软件的官网进行下载。

+ [VSCode](https://aka.ms/win32-x64-user-stable)
+ [Sublime](https://download.sublimetext.com/Sublime Text Build 3211 x64 Setup.exe)
+ [NotePad++](https://notepad-plus-plus.org/downloads/)

## 安装 Python3（必选）

在确认 Scoop 安装完成后，打开 PowerShell：

```powershell
Scoop install python
```

等待安装完成并尝试在 PowerShell 中输入:

```powershell
python
```

如果输出以下信息则代表安装正确：

```
Python 3.7.4 (tags/v3.7.4:e09359112e, Jul  8 2019, 20:34:20) [MSC v.1916 64 bit (AMD64)] on win32
Type "help", "copyright", "credits" or "license" for more information.
```

### 安装依赖（必选）

复制以下命令行至 PowerShell 或 CMD 或 bash 中，进行依赖的安装。

```powershell
# Python 中用于对 EXCEL 进行操作的组件
pip3 install xlwings
# Python 中用于访问 Window 剪贴板的组件
pip3 install pywin32
# Python 中用于传入参数的组件
pip3 install argparse
```

至此，程序所需运行的环境以及依赖都已准备完毕。

## 模板

模板的创建时为了让程序能够正常的读取并写入到实验条件数据库中，在此给出模板要遵循的格式：

```powershell
名称:值
```

例子：

```powershell
参数1:值1
参数2:值2
参数3:值3
参数4:值4
```

**请注意，如果你的实验记录里，标注了某个名称但是没有相对应的值，程序在接下来处理可能会出现错误。**


以上三个软件任选其一，但是我推荐首选用 VSCode，这三个软件的主要用途为编辑代码，这里不做过多展开。

## 程序

### 克隆仓库

将程序下载到本地，使用编辑器对程序进行自定义编辑，并对修改的内容依次测试确保能够正确使用。

~~此外，为了降低因文件路径而引起的 BUG，请将程序放置于实验记录Excel同一文件夹内。~~为了解决和避免这一情况的出现（如果变更文件路径则需要频繁的打开程序并修改，还有可能会引起不必要的 Bug），所以从新编写了程序，使其更为方便、与快捷。

现在，你不需要知道每一行代码的含义（事实上我也不知道，我估摸着能用），你只需要结合 Quicker 这个用于提升工作效率神器，将鼠标放在合适的位置并进行正确的点击即可。当然你仍然可以对程序进行个性化编辑，以满足自身的需求。

我将编写好的程序放在了 `extract_data`这个文件夹内，`idea`文件夹是我用于测试以及模板、测试数据的存放位置，`history`是存放之前用于解决问题的程序，用于存档和备忘。目前在`extract_data`文件夹内只有少部分程序。

未来会随着自身实验的推进和与师兄师姐的交流不断添加和优化新的程序以处理不同测试设备导出的数据、满足文件处理的要求。

### 运行程序

#### 通过命令行

虽然在这里介绍了如何通过命令行运行程序以得到自己期待的结果，但最终，我想大多数人都会采用另一种方式——通过 Quicker 调用程序。

再将仓库克隆到本地之后，如果需要依据自身的需要对程序进行个性化修改，请复制 `extract_data`文件夹至你熟悉的地方。为了能够在 PowerShell/CMD 中对程序进行调用，需要将放置程序的文件夹添加到系统的环境变量中，添加的方式如下：

1. 右键 **我的电脑**

2. 在弹出的窗口左侧找到 **高级系统设置**

3. 在弹出的窗口 **系统属性** 中，点击右下角的 **环境变量**

4. 在 环境变量 窗口中分为上下两个框，上方为用户变量，下方为系统变量。在 系统变量 中找到名为 **Path** 的变量，点击以选中 **Path**，然后点击 **编辑**，会弹出一个名为 **编辑环境变量** 的窗口，点击 **新建**，将存放程序的文件夹路径复制输入，随后 **确定→确定→确定**。

   请注意，你的环境变量格式应与下方内容类似：

   `F:\github_graduate\Valkyrie\extract_data`

5. 添加 环境变量 完成。

在存放程序的文件夹内按住 **shift** 键右击空白处，选择 **在此处打开 PowerShell/CMD 窗口**，在 PowerShell/CMD 中输入如下命令以查看如何使用：

```powershell
python Raman_Graphene.py -h
```

以执行`python Raman_Graphene.py  -h`为例，你将会在 PowerShell/CMD 中看到输出的帮助信息：

```powershell
usage: 该脚本旨在帮助你更方便的对拉曼数据进行整理 [-h] [-i] [-e] [-c] [-wc] [-wr] [-wf] [-cf] [-cr]

optional arguments:
  -h, --help         show this help message and exit
  -i , --input       要进行整理的拉曼数据文件路径
  -e , --excel       保存数据的 excel 文件路径

基础选项:
  -c , --column      要提取的数据列

进阶选项:
  -wc, --wconditon   模式：将实验条件写入指定 excel 表格
  -wr, --wraman      模式：将拉曼数据写入指定 excel 表格
  -wf, --wfit        模式：将拟合结果写入指定 excel 表格
  -cf, --cfit        模式：将拟合结果写入到剪贴板
  -cr, --copyselect  模式：将提取的数据列写入剪贴板
```

请按照输出的帮助信息对程序进行调用。请注意进阶选项中的模式不可同时使用，每次只能使用一个，如果同时使用将会按照最高优先级的那一个进行输出。

如果要修改程序，对程序配置完成后，请保存并进行测试，确保正常运行。

#### 通过Quicker

**相关程序的对应捷径可在本项目 Release 界面看到，或从以下网址查看分享的 Quicker 动作：**

捷径分享地址：https://getquicker.net/User/42/nekosan 

捷径的使用方式，请参阅此处： https://www.yuque.com/quicker/help/install-action 

## 解决方案

我对于这个项目希望它未来能做到全平台化(Win、Mac、Linux、IOS、Android)，能够集实验数据汇总、读取、云同步、案例对比和周汇报报告生成为一体。

目前因为个人能力有限，暂时只做到了实验数据的提取和写入，UI 界面会在以后部署（但不知道会在何时出现），不知道在研究生期间能不能完成，也许这个饼永远都不会完成。所以暂时就先采用 Quicker 进行调用。

我目前的解决方案：

+ 采用 ideaNote 进行实验记录的云同步
+ 使用 坚果云 同步所有的实验数据
+ Excel、Origin 进行数据分析

## 最后

如果这个小小的程序能够解决你的一部分痛点，目的也就达成了。

欢迎在 Github 上提供需求和 BUG 反馈和赞助。

***

## PDF文献小组知识共享方案（Mendeley）

1. 将所有文献拖进 Mendeley，进行整理。
2. 创建 Github 仓库或者云盘的同步文件夹
3. 打开Mendeley→Watched Folder→选择你存储文献的地方。
4. 打开Mendeley→ Tools→File Organizer→Browse→选择存储位置为本地
5. 复制整理好的文件夹至Github仓库或者云盘同步文件夹→上传
6. 打开Mendeley→右键你要读的文献→Open file Externally→打开文献→在需要的地方添加注释→保存并关闭
7. https://getquicker.net/sharedaction?code=c042c5b4-4e97-4e6f-96eb-08d737264562
8. 在浏览器打开上述网址，安装快捷动作
9. 单击以选中修改完的PDF文件，调出 Quicker 快捷面板，使用 **PDF注释提取快捷动作** 提取注释构建PDF注释全文检索数据库。
10. 复制整理好的文件夹至Github仓库或者云盘同步文件夹→上传