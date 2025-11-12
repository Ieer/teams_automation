# from src.teams_automation.client import TeamsClient
# python -m build --no-isolation

from teams_automation.client import TeamsClient

client = TeamsClient()


# 发送消息到指定的聊天窗口
# client.send_message(
# 	chat_name="teams ChatBot",
# 	message="Hello from TeamsClient!",
# 	section_name="Chats"
# )

# 发送文件
# file_path = r"D:\Python Robot\TeamsDemo\team_tree.png"  # 替换为实际的文件路径
# client.send_files(
# 	chat_name="teams ChatBot",
# 	filepaths=[file_path],
# 	section_name="Chats"
# )

# 发送图片和消息
client.send_message(
    "Hello, this is a test message from UIAutomation!",
	chat_name="teams ChatBot",
	section_name="Chats",
	image_path=r"D:\Python Robot\TeamsDemo\team_tree.png",
	close_filter=True
)