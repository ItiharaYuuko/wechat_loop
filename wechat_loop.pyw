import win32api, win32gui, win32con
import win32com.client as wclt
import win32clipboard as clipbd
import time, re, datetime
from tkinter import Label, Button, Tk, Entry, Frame, StringVar, IntVar, DoubleVar, Text, Spinbox, Checkbutton, BooleanVar, messagebox, HORIZONTAL
from tkinter.constants import END
import threading as thd
from pythoncom import CoInitialize

class GUI_MAIN_APP(Frame):
    def __init__(self, master = None) -> None:
        super().__init__(master)
        self.master = master
        self.pack()
        self.create_widget()

    def w_size(self, content: str) -> int:
        return len(content) * 12

    def create_widget(self) -> None:
        self.place(x = 0, y = 0, width=450, height=700)
        note = '''说明: 在输入框种输入要传输的内容选择文件复选框后,
        会在倒计时之后直接发送剪贴板中的内容, 默认的发送日期为当天日期,
        在日期输入框中输入日期格式例如2022-01-01, 
        时间格式例如13:45:23,7:23:05.'''
        self.note_label = Label(self, text = note, justify = 'left').place(x = 10, y = 0, width = 400, height = 65)

        l1_ct = '窗口名称：'
        Label(self, text=l1_ct).place(x = 10, y = 70, width = self.w_size(l1_ct))
        
        wl_ct = '文件传输助手'
        self.wlct = StringVar()
        self.wlct.set(wl_ct)
        self.window_name_entry = Entry(self, textvariable = self.wlct)
        self.window_name_entry.place(x = self.w_size(l1_ct) + 10, y = 70, width = 196)
        
        l2_ct = '要发送的内容: '
        Label(self, text = l2_ct).place(x = 10, y = 100, width = self.w_size(l2_ct) - 10)
        self.content_text = Text(self)
        self.content_text.place(x = 10, y = 130, width = 400, height = 300)

        l3_ct = '循环次数：'
        self.lts_ct = IntVar()
        Label(self, text = l3_ct).place(x = 10, y = 435, width = self.w_size(l3_ct))
        self.loop_times_spinbox = Spinbox(self, from_ = 1, to = 100000, textvariable = self.lts_ct)
        self.loop_times_spinbox.place(x = self.w_size(l3_ct) + 10, y = 435)

        l4_ct = '循环时间间隔(秒)：'
        Label(self, text = l4_ct).place(x = 10, y = 465, width = self.w_size(l4_ct) - 15)

        self.lti_ct = DoubleVar()
        self.lti_ct.set(1)
        self.loop_time_interval_entry = Entry(self, textvariable = self.lti_ct)
        self.loop_time_interval_entry.place(x = self.w_size(l4_ct) + 10, y = 465, width = 60)

        l5_ct = '日期：'
        Label(self, text = l5_ct).place(x = 10, y = 495, width = self.w_size(l5_ct))

        dt = time.strftime('%Y-%m-%d', time.localtime())
        self.dt_ct = StringVar()
        self.dt_ct.set(dt)
        self.send_date_entry = Entry(self, textvariable = self.dt_ct)
        self.send_date_entry.place(x = self.w_size(l5_ct) + 10, y = 495, width = self.w_size(dt))

        l6_ct = '时间：'
        Label(self, text = l6_ct).place(x = self.w_size(dt) + self.w_size(l5_ct) + 20,
                                        y = 495, width = self.w_size(l6_ct))

        tt = time.strftime('%H:%M:%S', time.localtime())
        self.tt_ct = StringVar()
        self.tt_ct.set(tt)
        self.send_time_entry = Entry(self, textvariable = self.tt_ct)
        self.send_time_entry.place(x = self.w_size(dt) + self.w_size(l5_ct) + self.w_size(l6_ct) + 30,
                                    y = 495, width = self.w_size(tt))

        l7_ct = '是否为文件(勾选后为True, 不勾选为False)'
        Label(self, text = l7_ct).place(x = 10, y = 525, width = round(self.w_size(l7_ct) * 0.72))

        self.fcbv = BooleanVar()
        self.file_check_box = Checkbutton(self, command = self.file_check_box_action)
        self.file_check_box.place(x = round(self.w_size(l7_ct) * 0.72) + 10, y = 525)

        self.indicator_content = StringVar()
        Label(self, textvariable = self.indicator_content).place(x = 150, y = 580, width = 120)
        
        self.send_button = Button(self, text = '发送', command = self.send_button_action)
        self.send_button.place(x = 155, y = 600, width = 100, height = 50)

    def file_check_box_action(self) -> None:
        if not self.fcbv.get():
            self.fcbv.set(not self.fcbv.get())
        else:
            self.fcbv.set(not self.fcbv.get())

    def send_button_action(self) -> None:
        send_date = self.dt_ct.get()
        send_time = self.tt_ct.get()
        rgx = re.compile(r'\d{4}\-\d{2}-\d{2} \d{1,2}:\d{2}:\d{2}')
        input_date_time = f'{send_date} {send_time}'
        if rgx.match(input_date_time):
            time_left = get_date_time_sub(send_time, send_date)
            if thd.active_count() < 2:
                loop_times = self.lts_ct.get()
                loop_time_interval = self.lti_ct.get()
                file_flag = self.fcbv.get()
                window_name = self.wlct.get()
                send_content = self.content_text.get(1.0, END)
                self.indicator_content.set(f'已经过：{0}%')
                msg = {'window_name': window_name, 'send_content': send_content,
                        'loop_times': loop_times, 'loop_time_interval': loop_time_interval,
                        'file_flag': file_flag, 'send_date': send_date, 'send_time': send_time,
                        'input_date_time': input_date_time}
                mt = thd.Thread(target = self.indicator_increse, args = (time_left, msg))
                mt.start()
        else:
            messagebox.showerror('时间日期错误', '请重新输入时间或日期再次尝试')

    def indicator_increse(self, time_left: int, msg: dict) -> None:
        for i in range(time_left + 1):
            self.indicator_content.set(f'已经过：{i / time_left * 100:.2f}%')
            time.sleep(1)
            self.update()
        messagebox.showinfo('发送提醒', '正在执行发送任务...')
        loop_execute(msg.get('window_name'), msg.get('send_content'), msg.get('loop_time_interval'),
                    msg.get('loop_times'), msg.get('file_flag'))
        messagebox.showinfo('任务提醒', '发送任务已完成。')

def get_window_handler(wname: str) -> int:
    win_ha = win32gui.FindWindow('ChatWnd', wname)
    win32gui.BringWindowToTop(win_ha)
    CoInitialize()
    sl = wclt.Dispatch('WScript.Shell')
    sl.SendKeys('%')
    win32gui.SetForegroundWindow(win_ha)
    return win_ha

def message_send(win_ha: int) -> None:
    win32api.keybd_event(17, 0, 0, 0)
    time.sleep(0.1)
    win32gui.SendMessage(win_ha, win32con.WM_KEYDOWN, 86, 0)
    time.sleep(0.1)
    win32gui.SendMessage(win_ha, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
    win32api.keybd_event(17, 0, win32con.KEYEVENTF_KEYUP, 0)

def content_copy_to_clipboard(text):
    clipbd.OpenClipboard()
    clipbd.EmptyClipboard()
    clipbd.SetClipboardText(text)
    clipbd.CloseClipboard()

def loop_execute(window_name: str,
                sending_text: str,
                time_interval: float,
                loop_times: int,
                file_flg: bool) -> None:
                
    for _ in range(1, loop_times + 1):
        time.sleep(time_interval)
        if not file_flg:
            content_copy_to_clipboard(sending_text)
        w_ha = get_window_handler(window_name)
        message_send(w_ha)

def get_file_flg() -> bool:
    cont = input('If file send <True> else <False> type here: ').strip()
    if cont == 'True':
        return True
    else:
        return False

def get_date_time_sub(send_time: str, send_date: str = datetime.datetime.today().date().strftime('%H:%M:%S')) -> int:
    ti = time.localtime()
    fdt = f'{send_date} {send_time}'
    td = time.strptime(fdt, '%Y-%m-%d %H:%M:%S')
    tx = datetime.datetime(ti.tm_year, ti.tm_mon, ti.tm_mday, ti.tm_hour, ti.tm_min, ti.tm_sec)
    ty = datetime.datetime(td.tm_year, td.tm_mon, td.tm_mday, td.tm_hour, td.tm_min, td.tm_sec)
    return (ty - tx).days * 24 * 3600 + (ty - tx).seconds

if __name__ == '__main__':
    root = Tk()
    root.title('微信轰炸工具 Ver GM 1.0.1')
    root.geometry('420x700+200+50')
    app = GUI_MAIN_APP(root)
    app.mainloop()
