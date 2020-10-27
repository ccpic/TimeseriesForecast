from pywinauto.application import Application
import time
from pywinauto import keyboard, mouse


USERNAME1 = 'chencheng'
PASSWORD1 = 'cc_137813'
USERNAME2 = 'chengtao'
PASSWORD2 = '888888'
# 对于Windows中自带应用程序，直接执行，对于外部应用应输入完整路径
app= Application(backend="uia").start(r'C:\Program Files (x86)\CRM系统管理\pmgr.exe')
time.sleep(2)

# 登录
keyboard.send_keys(USERNAME1)
keyboard.send_keys('{VK_TAB}')
keyboard.send_keys(PASSWORD1)
keyboard.send_keys('{VK_RETURN}')
time.sleep(10)
keyboard.send_keys(USERNAME2)
keyboard.send_keys('{VK_TAB}')
keyboard.send_keys(PASSWORD2)
keyboard.send_keys('{VK_RETURN}')

# 打开发货分析
mouse.click(coords=(208, 49))
mouse.click(coords=(265, 212))
mouse.click(coords=(300, 212))