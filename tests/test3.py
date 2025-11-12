import uiautomation as auto
import time
import os
import socket

from io import BytesIO
from PIL import Image
import win32clipboard as wc
import win32con

# 获取电脑名称
computer_name = socket.gethostname()
print(f">>> Computer Name: {computer_name}")
# 获取当前用户名称
current_user = os.getlogin()
print(f">>> Current User: {current_user}")


def send_message_in_teams(message):
	
	# 打开 Microsoft Teams 窗口，获取第一个 Teams 窗口
	windows_name = None
	for window in auto.GetRootControl().GetChildren():
		if "Microsoft Teams" in window.Name and "TeamsWebView" in window.ClassName:
			window.SetActive()  # 激活窗口
			print(f">>> Set active: {window.Name}")
			windows_name = window.Name
			break
	time.sleep(5) 

	# 获取 Microsoft Teams 窗口句柄
	teams_window = auto.WindowControl(searchDepth=1, Name=windows_name)
	if not teams_window.Exists(0, 0):
		raise Exception("Microsoft Teams is not running or the window is not found.")
	
	# 遍历 Teams 窗口的子控件，查找搜索按钮
	for child in teams_window.GetChildren():
		print(f"--> Child Name: {child.Name}, ControlType: {child.ControlTypeName}, ClassName: {child.ClassName}")
		
	# 切换到对话框
	chat_button = teams_window.Control(searchDepth=30, Name="Chat (Ctrl+2)")
	if not chat_button.Exists(0, 0):
		raise Exception("Chat button is not found. Make sure the Teams window is open and visible.")
	chat_button.Click() 
	teams_window.SetActive()  # 确保 Teams 窗口处于活动状态

	# 点击搜索按钮
	search_button = teams_window.ButtonControl(searchDepth=30, Name="Show filter text box (Ctrl+Shift+F)")
	# 判断存在并点击
	if not search_button.Exists(0, 0):
		raise Exception("Search button is not found. Make sure the Teams window is open and visible.")
	search_button.Click()

	# ! 输入搜索内容 
	group_name = "team ChatBot"  # 替换为实际的群组名称
	search_input = teams_window.EditControl(searchDepth=30, Name="Filter by name or group name")
	if not search_input.Exists(0, 0):
		raise Exception("Search input is not found. Make sure the Teams window is open and visible.")
	search_input.SetFocus()
	search_input.SendKeys(group_name)
	time.sleep(1)

	# 进入对话框
	chat_item = teams_window.Control(searchDepth=30, Name="Filter active")
	for item in chat_item.GetChildren():
		print(f"--> Chat Item Name: {item.Name}, ControlType: {item.ControlTypeName}")

	chat_group =['Favorites','Chats']
	item_object = [ item for item in chat_item.GetChildren() if item.Name in chat_group]
	
	print(f">>> Found chat items: {[item.Name for item in item_object]}")
 
	# 获取item_object的元素所有元素
	if not item_object:
		raise Exception("Chat group is not found. Make sure the Teams window is open and visible.")
	# 遍历item_object所有元素集合
	chat_all = {}
	for chat in item_object:
		print(f">>> Entering chat: {chat.Name}")
		# 继续遍历子元素，如果有子元素，添加到chat_all
		for sub_chat in chat.GetChildren():
			# 判断 sub_chat对象为空
			for sub in sub_chat.GetChildren():
				sub_name = sub.Name
				if sub_name != "":
					import re

					# 预编译正则、合并清理步骤、处理重复名称并去除不可见空白
					last_msg_re = re.compile(r"Last message.*", re.IGNORECASE)
					leading_re = re.compile(r'^(?:Group|Chat|chat)\b[:\-\u2013\u2014]?\s*', re.IGNORECASE)
					sub_group_name = last_msg_re.sub("", sub_name).strip()
					sub_group_name = leading_re.sub("", sub_group_name)
					# 规范化空白（包含不间断空格）并去除两端空格
					sub_group_name = re.sub(r'[\s\u00A0]+', ' ', sub_group_name).strip()
					# 特殊Group处理
					if "Teams Chatbot Bot" in sub_group_name:
						sub_group_name = "Columbus Teams Chatbot"

					if not sub_group_name:
						continue

					chat_all[f"{chat.Name}_{sub_group_name}"] = sub
					# print(f"----> Sub Chat: {sub_group_name}")
    
	print(chat_all)

	# 选择聊天（点击聊天组，规则：选择包含关键词 "Chats" 的聊天组）
	select_chat = "Chats"
	# 选择第一个匹配的聊天组键的值
	target_chat = next((value for key, value in chat_all.items() if select_chat in key), None)
	if target_chat is None:
		raise Exception(f"No chat found containing '{select_chat}'.")
	print(f">>> Clicking chat: {target_chat.Name}")
	target_chat.Click()
	time.sleep(3)

	# 关闭搜索框
	close_search = teams_window.ButtonControl(searchDepth=30, Name="Close filter text box")
	if close_search.Exists(0, 0):
		close_search.Click()
		time.sleep(1)

	
	# 输入文本消息内容
	message_input = teams_window.EditControl(searchDepth=30, Name="Type a message")
	if not message_input.Exists(0, 0):
		raise Exception("Message input box is not found. Make sure the Teams window is open and visible.")
	message_input.SetFocus()
	# 清空输入框
	message_input.SendKeys("{Ctrl}a{Del}")
	message_input.SendKeys(message)
	time.sleep(1)

	# 发送图片
	image_path = r"D:\Python Robot\TeamsDemo\team_tree.png"
	if os.path.exists(image_path):
		try:
			def _set_clipboard_image(img_path):
				with Image.open(img_path) as img:
					with BytesIO() as output:
						img.convert("RGB").save(output, format="BMP")
						bytes_data = output.getvalue()[14:]
				wc.OpenClipboard()
				wc.EmptyClipboard()
				wc.SetClipboardData(win32con.CF_DIB, bytes_data)
				wc.CloseClipboard()

			_set_clipboard_image(image_path)
			message_input.SetFocus()
			message_input.SendKeys("{Ctrl}v")
			time.sleep(2)
		except Exception as err:
			print(f">>> Failed to send image: {err}")
	else:
		print(f">>> Image not found: {image_path}")


	# 点击发送按钮
	send_button = teams_window.ButtonControl(searchDepth=30, Name="Send (Ctrl+Enter)")
	if not send_button.Exists(0, 0):
		raise Exception("Send button is not found. Make sure the Teams window is open and visible.")
	send_button.Click()
	print(">>> Message sent successfully.")
	# 等待几秒钟以确保消息发送完成
	time.sleep(5)
 
    # 将窗口最小化
	auto.SendKeys("{win}d")  # 最小化所有窗口



if __name__ == "__main__":
	# 示例消息
	message = "Hello, this is a test message from UIAutomation!"
	
	# 调用函数发送消息
	send_message_in_teams(message)


