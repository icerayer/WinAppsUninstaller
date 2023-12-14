import subprocess
import PySimpleGUI as sg
import pyperclip


# 版本变更：0.8.1
# 修改tidy()，绿色放前面不变，黑色放中间，紫色放到最后

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
                    print(app, '确认不存在')
                    window[app].update(text_color='red')
                else:
                    print('应用还存在，似乎未卸载成功')
        except Exception as e:
            print(f"卸载应用时出现错误： {e}")
        print('\n')


def notfoundAppx(app_name):
    command = f"Get-AppxPackage -Name '*{app_name}*'"
    try:
        check_result = subprocess.run(["powershell", command], capture_output=True, text=True)
        if check_result.stdout.strip() == "":
            return True
        else:
            return False
    except Exception as e:
        print(f"An error occurred: {e}")


def getColor(curr_name):
    if curr_name in warn_list:
        return 'purple'
    elif curr_name in recom_dict.keys():
        return 'green'
    else:
        return 'black'


def getName(name):
    if name in recom_dict.keys():
        return recom_dict[name]
    else:
        return name


def tidy(name_list):
    top = []
    tail = []
    mid = []
    for x in name_list:
        if x in recom_dict.keys():
            top.append(x)
        elif x in warn_list:
            tail.append(x)
        else:
            mid.append(x)
    top.sort()
    tail.sort()
    print('推荐卸载', len(top))
    return top + mid + tail


def create_layout():
    layout = []
    col_num = 4
    row_num, remainder = divmod(len(name_list), col_num)
    if remainder:
        row_num += 1
    start = 0
    for row in range(row_num):
        row_lay = []
        if remainder and row == row_num - 1:
            col_num = remainder
        for flag in range(col_num):
            curr_name = name_list[start + flag]
            unit_lay = [
                sg.Checkbox(getName(curr_name), size=(30, 1), key=curr_name, text_color=getColor(curr_name)),
                sg.Button('卸载', key=f'button-{curr_name}')
            ]
            row_lay.append(sg.Column([unit_lay]))
        layout.append(row_lay)
        start += col_num

    layout.append([
        sg.Button('全选'),
        sg.Button('反选'),
        sg.Button('复制名称'),
        sg.Button('一键卸载'),
        sg.Button('退出'),
        sg.Text('推荐卸载', text_color='green'),
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
