import subprocess
import PySimpleGUI as sg
import pyperclip


# 版本变更：0.8.2
# 增加第三方分类Developer

# appinfdic的key有：["Name", "Publisher", "Architecture", "ResourceId", "Version","PackageFullName", "InstallLocation", "IsFramework", "PackageFamilyName","PublisherId", "IsResourcePackage", "IsBundle", "IsDevelopmentMode","NonRemovable", "IsPartiallyStaged", "SignatureKind", "Status", 'Dependencies']
def uninstall_selected_apps(sel_apps):
    for app in sel_apps:
        try:
            if notfoundAppx(app):
                print(f"应用 '{app}' 不存在")
            else:
                print(f"应用 '{app}' 存在", "开始卸载...")
                uninstall_command = f"Get-AppxPackage -Name '*{app}*' | Remove-AppxPackage"
                result = subprocess.run(["powershell", uninstall_command], capture_output=True, text=True)
                if result.returncode == 0:
                    print("卸载命令执行成功！")
                else:
                    print('命令执行失败，参考：')
                    print(result.stdout)
                    print(result.stderr)
                if notfoundAppx(app):
                    print(app, '确认已卸载!')
                    window[app].update(text_color='red')
                else:
                    print('应用还存在，似乎未卸载成功')
        except Exception as e:
            print(f"卸载应用时出现错误： {e}")
        print('\n')


def notfoundAppx(app_name):  # 检查指定的应用程序是否不存在
    # 定义一个命令，用于搜索指定应用程序
    command = f"Get-AppxPackage -Name '*{app_name}*'"
    try:
        # 使用PowerShell运行命令，并将输出捕获到check_result中
        check_result = subprocess.run(["powershell", command], capture_output=True, text=True)
        # 如果输出为空，则表示应用程序不存在
        if check_result.stdout.strip() == "":
            return True
        else:
            return False
    except Exception as e:
        # 如果出现异常，则打印异常信息
        print(f"An error occurred: {e}")


def getColor(curr_name):
    if curr_name in warn_list:
        return 'purple'
    elif curr_name in recom_dict.keys():
        return 'green'
    elif curr_name in dev_list:
        return 'blue'
    else:
        return 'black'


def getName(name):  # 获取推荐列表中的名字
    # 判断name是否在recom_dict中
    if name in recom_dict.keys():
        # 如果在，返回recom_dict中的值
        return recom_dict[name]
    else:
        # 如果不在，返回name
        return name


def tidy(name_list):
    recom = []
    noml = []
    dev = []
    warn = []

    for x in name_list:
        if x in recom_dict.keys():
            recom.append(x)
        elif x in warn_list:
            warn.append(x)
        elif x in dev_list:
            dev.append(x)
        else:
            noml.append(x)
    recom.sort()
    noml.sort()
    dev.sort()
    warn.sort()
    print('推荐卸载\t', len(recom))
    print('第三方\t', len(dev))
    print('不建议卸载\t', len(warn))
    return recom + noml + dev + warn


def create_layout():
    layout = []
    col_num = 4  # 列数为4，一行显示4个
    row_num, remainder = divmod(len(name_list), col_num)
    if remainder:  # 如果还有剩余，则行数加1
        row_num += 1
    start = 0  # 每行起始元素在name_list中的位置索引
    for row in range(row_num):  # 遍历行
        row_lay = []
        if remainder and row == row_num - 1:  # 如果是最后一行
            col_num = remainder  # 列数等于取模剩余的个数
        for flag in range(col_num):  # 遍历每行中的元素位置索引
            curr_name = name_list[start + flag]  # 根据当前元素在name_list中的位置索引(起始位索引+行内索引)获取应用名
            unit_lay = [  # 绘制每个应用的布局
                sg.Checkbox(getName(curr_name), size=(30, 1), key=curr_name, text_color=getColor(curr_name)),
                sg.Button('卸载', key=f'button-{curr_name}')
            ]
            row_lay.append(sg.Column([unit_lay]))  # 将每个应用的布局添加进row布局
        layout.append(row_lay)  # 把每行添加进layout
        start += col_num  # 每行起始元素在name_list中的位置索引+列数，等于下一行首位的索引

    layout.append([
        sg.Button('全选'),
        sg.Button('反选'),
        sg.Button('复制名称'),
        sg.Button('一键卸载'),
        sg.Button('退出'),
        sg.Text('推荐卸载', text_color='green'),
        sg.Text('第三方', text_color='blue'),
        sg.Text('不推荐卸载', text_color='purple'),
        sg.Text('已卸载', text_color='red'),
    ])
    return layout


recom_dict = {
    "Microsoft.Print3D": "打印3D",
    "Microsoft.WindowsCamera": "相机",
    "Microsoft.GetHelp": "获取帮助",
    "Microsoft.WindowsFeedbackHub": "Windows反馈中心",
    "Microsoft.Getstarted": "入门提示",
    "Microsoft.MicrosoftOfficeHub": "Office Hub",
    "Microsoft.ZuneMusic": "Zune音乐",
    "Microsoft.ZuneVideo": "Zune视频",
    "Microsoft.Microsoft3DViewer": "3D查看器",
    "Microsoft.MSPaint": "画图",
    "Microsoft.Paint": "画图",
    "Microsoft.MicrosoftStickyNotes": "便签",
    "Microsoft.Wallet": "钱包",
    "Microsoft.WindowsMaps": "地图",
    "Microsoft.Messaging": "消息",
    "Microsoft.BingWeather": "天气",
    "Microsoft.People": "人脉",
    "Microsoft.YourPhone": "你的手机",
    "Microsoft.OneConnect": "OneConnect",
    "Microsoft.MixedReality.Portal": "混合现实门户",
    "Microsoft.Xbox.TCUI": "Xbox TCUI",
    "Microsoft.XboxApp": "XboxApp",
    "Microsoft.XboxGameOverlay": "XboxGameOverlay",
    "Microsoft.XboxGamingOverlay": "XboxGamingOverlay",
    "Microsoft.XboxSpeechToTextOverlay": "Xbox语音转文本叠加层",
    "Microsoft.XboxIdentityProvider": "Xbox身份提供者",
    "Microsoft.GamingApp": "GamingApp",
    "Microsoft.3DBuilder": "3D建模",
    "Microsoft.MicrosoftSolitaireCollection": "微软纸牌游戏合集",
    "Microsoft.Office.OneNote": "OneNote",
    "Microsoft.SkypeApp": "Skype应用",
    "Microsoft.Windows.Photos": "照片",
    "Microsoft.WindowsAlarms": "闹钟",
    "Microsoft.windowscommunicationsapps": "邮件和日历",
    "microsoft.windowscommunicationsapps": "邮件和日历",
    "Microsoft.MicrosoftEdge.Stable": "Microsoft Edge稳定版",
    "Microsoft.Todos": "Todos待办事项",
    "Microsoft.BingNews": "资讯",
    "Clipchamp.Clipchamp": "Clipchamp视频编辑器",
    "Microsoft.549981C3F5F10": "Cortana",
    "Microsoft.WindowsCalculator": "计算器",
    "Microsoft.WindowsSoundRecorder": "录音机"
}
warn_list = ['Microsoft.WindowsStore', 'Microsoft.WindowsTerminal', 'Microsoft.DesktopApplnstaller']

powershell_command = "Get-AppxPackage"
echo = subprocess.check_output(["powershell", powershell_command]).decode("utf-8")
echo_list = list(filter(None, echo.split("\r\n\r\n")))

# 创建一个空列表，用于存储包的信息
pkg_list = []  # 包列表
name_list = []
dev_list = []
# 遍历echo_list中的每个元素（即每个包的信息）
for echo_item in echo_list:
    # 将字符串按照换行符进行分割，得到一个列表
    lines = str(echo_item).split("\n")
    # 创建一个空字典，用于存储当前包的信息
    appinfdic = {}
    # 遍历lines中的每个元素（即当前包的一行信息）
    for x in lines:
        # 如果当前行包含冒号，说明该行是一个键值对
        if ":" in x:
            # 使用split方法将字符串按照冒号进行分割，得到一个列表，其中第一个元素为键，第二个元素为值
            k, v = x.split(":", 1)
            # 将键和值去除首尾空格后存入字典
            appinfdic[k.strip()] = v.strip()
    # 只添加"NonRemovable"为"False"的包到包列表中
    if appinfdic.get("NonRemovable") == "False" and appinfdic.get("IsFramework") == "False":
        pkg_list.append(appinfdic)  # 将一个包字典存入包列表
        name_list.append(appinfdic.get('Name'))
    if appinfdic.get("SignatureKind") == "Developer":
        dev_list.append(appinfdic.get('Name'))

name_list = tidy(name_list)  # 推荐列表中的应用置顶
print("可卸载的非框架应用", len(name_list))
# for appx in name_list:
#     print(appx)

# 启动窗体
sg.theme('LightGrey1')
custom_font = ("Microsoft YaHei", 9)
window = sg.Window('可卸载的非框架应用', create_layout(), font=custom_font)

while True:
    event, values = window.read()

    if event in (sg.WIN_CLOSED, '退出'):
        break

    if event.startswith('button-'):
        app_name = [event.split('-')[1]]
        uninstall_selected_apps(app_name)

    if event == '全选':
        for app in name_list:
            window[app].update(value=True)

    if event == '反选':
        for app in name_list:
            window[app].update(value=not window[app].get())

    if event == '一键卸载':
        selected_apps = [app for app in name_list if window[app].get()]
        uninstall_selected_apps(selected_apps)

    if event == '复制名称':
        selected_apps = [app for app in name_list if window[app].get()]
        print(selected_apps)
        pyperclip.copy('\n'.join(selected_apps))

window.close()
