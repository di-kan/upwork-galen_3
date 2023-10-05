import wx
from enum import Enum
from threading import Thread
import subprocess
import shutil
import os
import wx.lib.intctrl
import winreg
import platform


class States(Enum):
    # When program is run
    DEFAULT = 1
    # When excel has been opened but browsers are not
    IDLE = 2
    # When excel is open and browsers are as well
    IDLE_BROWSERS = 3
    # When it has started processing
    CRAWLING = 4

class MainWindow(wx.Frame):
    def __init__(self, parent, title, eng):
        def _init_controls(panel, border):
            bor_h = border[0]
            bor_v = border[1]
            # 5 row sizers
            row1_sizer = wx.BoxSizer(wx.HORIZONTAL)
            row2_sizer = wx.BoxSizer(wx.HORIZONTAL)
            row3_sizer = wx.BoxSizer(wx.HORIZONTAL)
            row4_sizer = wx.BoxSizer(wx.HORIZONTAL)
            row5_sizer = wx.BoxSizer(wx.HORIZONTAL)
            row6_sizer = wx.BoxSizer(wx.HORIZONTAL)


            # CREATE CONTROLS TO PANEL
            self.btn_xls_open = wx.Button(panel, wx.ID_ANY, "Open")
            self.lbl_xls = wx.StaticText(panel, label=f"")
            self.lbl_from = wx.StaticText(panel, label=f"Process rows from:")
            self.txt_from = wx.lib.intctrl.IntCtrl(panel, value=0, size=txt_size)
            self.lbl_to = wx.StaticText(panel, label=f" up to:")
            self.txt_to = wx.lib.intctrl.IntCtrl(panel, value=0, size=txt_size)
            self.lbl_threads = wx.StaticText(panel, label=f"Threads:")
            self.cmb_threads = wx.ComboBox(panel, value="1",
                                                choices=["1","2","3","4","5","6","7","8"],
                                                style=wx.CB_READONLY)
            self.btn_launch_browsers = wx.Button(panel, wx.ID_ANY, "Launch Browsers")
            self.btn_start = wx.Button(panel, wx.ID_ANY, "Start")
            self.btn_stop = wx.Button(panel, wx.ID_ANY, "Stop")
            self.gauge = wx.Gauge(panel, 100, size=(350, 20))
            self.statusBar = self.CreateStatusBar(style=wx.BORDER_NONE)
            self.statusBar.SetStatusText("Open an excel file")

            #ASSIGN FUNCTIONS TO BUTTONS
            self.Bind(wx.EVT_BUTTON, self.open_excel, self.btn_xls_open)
            self.Bind(wx.EVT_BUTTON, self.launch_browsers, self.btn_launch_browsers)
            self.Bind(wx.EVT_BUTTON, self.start, self.btn_start)
            self.Bind(wx.EVT_BUTTON, self.stop, self.btn_stop)

            # ADD CONTROLS TO EACH ROW
            row1_sizer.Add(self.btn_xls_open, 1, wx.LEFT | wx.RIGHT, bor_h)
            row2_sizer.Add(self.lbl_xls, 1, wx.LEFT | wx.RIGHT, bor_h)

            row3_sizer.Add(self.lbl_from, 1, wx.LEFT | wx.RIGHT | wx.ALIGN_CENTER, bor_h)
            row3_sizer.Add(self.txt_from, 2, wx.LEFT | wx.RIGHT | wx.ALIGN_CENTER, bor_h)
            row3_sizer.Add(self.lbl_to, 1, wx.LEFT | wx.RIGHT | wx.ALIGN_CENTER, bor_h)
            row3_sizer.Add(self.txt_to, 2, wx.LEFT | wx.RIGHT | wx.ALIGN_CENTER, bor_h)

            row4_sizer.Add(self.lbl_threads, 1, wx.LEFT | wx.RIGHT | wx.ALIGN_CENTER, bor_h)
            row4_sizer.Add(self.cmb_threads, 1, wx.LEFT | wx.RIGHT | wx.ALIGN_CENTER, bor_h)
            row4_sizer.Add(self.btn_launch_browsers, 2, wx.LEFT | wx.RIGHT | wx.ALIGN_CENTER, bor_h)

            row5_sizer.Add(self.btn_start, 1, wx.LEFT | wx.RIGHT, bor_h)
            row5_sizer.Add(self.btn_stop, 1, wx.LEFT | wx.RIGHT, bor_h)
            row6_sizer.Add(self.gauge, 1, wx.LEFT | wx.RIGHT, bor_h)

            return [row1_sizer, row2_sizer, row3_sizer, row4_sizer, row5_sizer, row6_sizer]

        wx.Frame.__init__(self, parent, title=title)
        panel = wx.Panel(self)
        match platform.system().lower():
            case "linux":
                self.SetMinSize((600, 450))
                txt_size = 100, 40
                borders = (5, 5)
            case "windows":
                self.SetMinSize((600, 450))
                borders = (5, 1)
                txt_size = 100, 20

        self.eng = eng
        self.ports = []
        self.browsers_processes = []
        self.state = States.DEFAULT
        # MAIN APP SIZER
        col_box = wx.BoxSizer(wx.VERTICAL)
        #CREATE BOX SIZERS
        rows = _init_controls(panel, border=borders)
        # ADD BOX_SIZERS TO MAIN APP SIZER
        for row in rows:
            col_box.Add(row, 1, wx.ALL | wx.EXPAND, border=10)

        # RESIZE AND SHOW
        panel.SetSizer(col_box)
        self.Centre()

        self.timer = wx.Timer(self)
        self.Bind(wx.EVT_TIMER, self.update_gui, self.timer)
        self.Bind(wx.EVT_CLOSE, self.on_close_window)
        self.set_gui_state(States.DEFAULT)
        self.timer.Start(500)
        self.Show()

    def on_close_window(self, event):
        self.kill_browsers()
        self.Destroy()

    def update_gui(self, event):
        if self.state == States.CRAWLING:
            if self.eng.current > 0:
                self.gauge.SetValue(self.eng.current)
                self.statusBar.SetStatusText(f"Company {self.eng.current}/{self.eng.processing_totals}")
        else:
            self.gauge.SetValue(0)

    def set_gui_state(self, state):
        self.state = state
        self.statusBar.Enable(True)
        wx.Yield()
        match state:
            case States.DEFAULT:
                self.btn_xls_open.Enable(True)
                self.lbl_from.Enable(False)
                self.txt_from.Enable(False)
                self.lbl_to.Enable(False)
                self.txt_to.Enable(False)
                self.lbl_threads.Enable(False)
                self.cmb_threads.Enable(False)
                self.btn_launch_browsers.Enable(False)
                self.btn_start.Enable(False)
                self.btn_stop.Enable(False)
                self.gauge.Enable(False)
                self.statusBar.SetStatusText("Open an excel file")
            case States.IDLE:
                self.btn_xls_open.Enable(False)
                self.lbl_from.Enable(True)
                self.txt_from.Enable(True)
                self.lbl_to.Enable(True)
                self.txt_to.Enable(True)
                self.lbl_threads.Enable(True)
                self.cmb_threads.Enable(True)
                self.btn_launch_browsers.Enable(True)
                self.btn_start.Enable(False)
                self.btn_stop.Enable(False)
                self.gauge.Enable(True)
                self.statusBar.SetStatusText("Launch Browsers")
            case States.IDLE_BROWSERS:
                self.btn_xls_open.Enable(False)
                self.lbl_from.Enable(True)
                self.txt_from.Enable(True)
                self.lbl_to.Enable(True)
                self.txt_to.Enable(True)
                self.lbl_threads.Enable(True)
                self.cmb_threads.Enable(False)
                self.btn_launch_browsers.Enable(False)
                self.btn_start.Enable(True)
                self.btn_stop.Enable(False)
                self.gauge.Enable(True)
                self.statusBar.SetStatusText("Bypass any checks in the browsers if any and start")
            case States.CRAWLING:
                self.btn_xls_open.Enable(False)
                self.lbl_from.Enable(True)
                self.txt_from.Enable(True)
                self.lbl_to.Enable(True)
                self.txt_to.Enable(True)
                self.lbl_threads.Enable(True)
                self.cmb_threads.Enable(False)
                self.btn_launch_browsers.Enable(False)
                self.btn_start.Enable(False)
                self.btn_stop.Enable(True)
                self.gauge.Enable(True)
        wx.Yield()

    def open_excel(self, event):
        with wx.FileDialog(self,
                "Open a excel file",
                defaultDir = "",
                defaultFile = "",
                wildcard = "Excel files (*.xls, *.xlsx)|*.xls;*.xlsx",
                style = wx.FD_OPEN) as dlg:
            if dlg.ShowModal() == wx.ID_OK:
                complete_filename = dlg.GetPaths()[0]
                self.set_gui_state("open")
                self.eng.excel_filename = complete_filename
                self.txt_from.SetValue(2)
                self.txt_to.SetValue(self.eng.original_df.shape[0])
                self.lbl_xls.SetLabel(complete_filename)
                self.set_gui_state(States.IDLE)

    def launch_browsers(self, event):
        self.kill_browsers()
        chrome_path = self.get_chrome_path()
        starting_port = 9400
        current_dir = os.getcwd()
        profiles_path = os.path.join(current_dir, "profiles")
        if not os.path.exists(profiles_path):
            os.makedirs(profiles_path)
        for file in os.listdir(profiles_path):
            final_dir = os.path.join(profiles_path, file)
            if os.path.exists(final_dir):
                shutil.rmtree(final_dir)
        self.ports.clear()
        for i in range(0, int(self.cmb_threads.GetValue())):
            final_dir = os.path.join(profiles_path,str(i))
            port = starting_port+i
            self.ports.append(port)

            cmd = [chrome_path, f"--remote-debugging-port={port}", f"--user-data-dir={final_dir}", self.eng.url]
            p = subprocess.Popen(cmd)
            self.browsers_processes.append(p)
        self.set_gui_state(States.IDLE_BROWSERS)

    def start(self, event):
        start = int(self.txt_from.GetValue())
        stop = int(self.txt_to.GetValue())
        self.gauge.SetRange(stop-start+1)
        process_thread = Thread(target=self.eng.process_companies, args=(start, stop, self.ports))
        process_thread.daemon = True
        process_thread.start()
        self.set_gui_state(States.CRAWLING)
        # process_thread.join()
        # self.eng.process_companies(start, stop, self.ports)

    def stop(self, event):
        self.eng.working = False
        self.kill_browsers()
        self.set_gui_state(States.IDLE)

    def stop_from_engine(self):
        print("Got stop from engine")
        self.set_gui_state(States.IDLE)
        self.eng.working = False
        self.kill_browsers()

    def kill_browsers(self):
        for p in self.browsers_processes:
            print(f"killing {p.pid}")
            p.terminate()
        self.browsers_processes.clear()
    
    def get_chrome_path(self):
        result = None
        key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe")        
        chrome_path, _ = winreg.QueryValueEx(key, None)
        if os.path.exists(chrome_path):
            result = chrome_path
        return result
