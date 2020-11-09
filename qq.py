import win32gui,win32con
import win32clipboard as w
import time

#将消息写入剪贴板
def setText(text):
    w.OpenClipboard()
    w.EmptyClipboard()
    w.SetClipboardData(win32con.CF_UNICODETEXT, text)
    w.CloseClipboard()

#qq搜索栏搜索指定好友
def searchUser(name):
    #鼠标定位qq搜索栏
    hand = win32gui.FindWindow('TXGuiFoundation', 'QQ')
    setText(name)
    win32gui.SendMessage(hand, 770, 0, 0)
    #表示停止1.5秒再运行（运行太快qq会反应不过来）
    time.sleep(1.5)
    win32gui.SendMessage(hand, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)


#按重复次数发送消息
def formal():
    name = input("请输入联系人：")
    msg = input("请输入消息内容：")
    n = int(input("请输入重复消息次数："))
    t = float(input("请输入消息间隔时间/s："))
    def sendMsger1(name):
        #自动定位聊天窗口
        hand = win32gui.FindWindow('TXGuiFoundation', name)
        setText(msg)
        #重复发送消息
        for i in range(1,n+1):
            win32gui.SendMessage(hand,770,0,0)
            win32gui.SendMessage(hand,win32con.WM_KEYDOWN,win32con.VK_RETURN,0)
            i = i+1
            time.sleep(t)
        print("运行完成!")

    searchUser(name)
    time.sleep(1)
    print("开始发送")
    print('...')
    sendMsger1(name)

#按停止时间发送消息
def stoptimer():
    name = input("请输入联系人：")
    msg = input("请输入消息内容：")
    stoptime = input("请输入停止发送时间(2020-03-11 17:31:09)：")
    t = float(input("请输入消息间隔时间/s："))
    #转换为时间戳
    stoptime = time.mktime(time.strptime(stoptime,"%Y-%m-%d %H:%M:%S"))
    def sendMsger2(name):
        hand = win32gui.FindWindow('TXGuiFoundation', name)
        setText(msg)
        while True:
            #获取当前时间的时间戳
            nowtime = time.time()
            if nowtime < stoptime:
                win32gui.SendMessage(hand,770,0,0)
                win32gui.SendMessage(hand,win32con.WM_KEYDOWN,win32con.VK_RETURN,0)
                #可调
                time.sleep(t)
            elif nowtime > stoptime:
                break
        print("运行完成!")

    searchUser(name)
    time.sleep(1)
    print("开始发送")
    print('...')
    sendMsger2(name)


if __name__ == "__main__":
    print("使用说明：\n1.输入使用方法之前应先点击一下好友搜索栏\n2.输入使用方法时应输入1或者2\n3.发送消息时鼠标点击最小化即可停止发送，重新点击消息发送栏继续发送消息\n4.ctrl+c终止程序")
    a = input("\n请输入使用方法: \n1 手动输入次数\n2 输入时间自动运行\n")
    if a == '1':
        formal()
    elif a == '2':
        stoptimer()
