"""     
    ExcelVBA:
    ```
        Option Explicit
        Sub Run_Python()
            RunPython ("import ExcelController;ExcelController.main()")
        End Sub
    ```
    select -> RunPython
"""

import xlwings as xw
import glob as g
import pandas as pd
from natsort import natsorted
from practical_package import module_gui_text as mgt

class ProcessingTarget:
    def __init__(self, self_root, label_x, label_y, bar_x, bar_y, bar_len):
        self.flag_running = False
        self.label_progress = mgt.ProgressLabel(label_x, label_y)
        self.progressbar = mgt.Progressbar(self_root, bar_x, bar_y, bar_len)
        
    def target(self):
        self.flag_running = True

        # connect
        wb = xw.Book.caller()
        ws = wb.sheets('[sheet name]')

        # file analysis
        self.label_progress.update("Path analyzing...")
        paths = []  
        excs = []     # exclusive
        s = ws.range('[cel]').value
        if ";" in s:
            s = s.split(";")
            for i in s:
                if i != "":
                    paths.append(i)
        else:
            paths.append(s)
        s = ws.range('[cel]').value
        if ";" in s:
            s = s.split(";")
            for i in s:
                if i != "":
                    excs.append(i)
        else:
            excs.append(s)

        # match paths
        flag_not = False
        for p in paths:
            if not mgt.path_search_continue(p).is_ok:
                flag_not = True
                self.label_progress.update("The target path is incorrect. Please try again...")
                mgt.time.sleep(1)
                break
        if not flag_not:
            for e in excs:
                if not mgt.path_search_continue(e).is_ok:
                    flag_not = True
                    self.label_progress.update("The exclusive path is incorrect. Please try again...")
                    mgt.time.sleep(1)
                    break

        if not flag_not:
            n = 11  # start rows
            data_group = []
            while ws.range('[cel char1]' + str(n)).value is not None:
                self.label_progress.update("Loading..." + ws.range('[cel char2]' + str(n)).value)
                data_group.append(ws.range('[cel char2]' + str(n)).value)
                n += 1
            data_check = ws.range('[cel char3]11:[cel char4]' + str(n)).value

            # delete
            ws.range('[cel char4]11:[cel char5]' + str(n)).value = ""

            # file search
            self.label_progress.update("File analyzing...")
            list_file = []
            for p in paths:
                if p[0:3] != "\\/*":  # avoid deadlock
                    files = g.glob("{0}{1}".format("**/" if p == "\\" else p[1:] + "/**/" if p[0] == "\\" else p + "/**/", "*.csv"), recursive=True)
                    for f in files:
                        list_file.append(f)

            # exclusive processing
            cnt_element = mgt.counter()
            for _ in range(len(list_file)):
                if len(list_file) == cnt_element.result():
                    break
                flag_break = False
                for j in excs:
                    if j == list_file[cnt_element.result()][:len(j)]:
                        del list_file[cnt_element.result()]
                        flag_break = True
                        break
                if not flag_break:
                    cnt_element.count()

            # int regular expression sort
            list_file = natsorted(list_file)

            # group analysis
            df = [[] for _ in range(2)]   # group cols
            for file in list_file:
                order = mgt.lines_list("[group]", file)
                for line in order.line:
                    cnt_array = mgt.counter()
                    df[cnt_array.result()].append(file)
                    df[cnt_array.count()].append(line[len("[group]"):])

            # in database
            cnt_array = mgt.counter()
            df = pd.DataFrame({
                'file' : df[cnt_array.result()], 
                'group': df[cnt_array.count()]
                })

            # output
            cache = ""
            self.label_progress.update("Outputting...(Never close during output)")
            self.progressbar.set.configure(maximum=len(df))
            for index_df, item_df in df.iterrows():
                self.progressbar.update(index_df)
                if cache != item_df.file:
                    ws.range('[cel char4]' + str(index_df + 11)).value = mgt.os.path.basename(item_df.file)
                    cache = item_df.file
                ws.range('[cel char2]' + str(index_df + 11)).value = item_df.group
                ws.range('[cel char3]' + str(index_df + 11)).value = ["" for _ in range(7)]   # max cols
                for i in range(len(data_group)):
                    if data_group[i] == item_df.group:
                        ws.range('[cel char3]' + str(index_df + 11)).value = data_check[i]
                        break
        # finish
        self.flag_running = False
        self.label_progress.end("", flag_dt=True, flag_timer=True)
        
    def start(self):
        self.thread_target = mgt.threading.Thread(target = self.target)
        self.thread_target.setDaemon(True)
        self.thread_target.start()

class GuiApplication(mgt.tk.Frame):
    def __init__(self, master=None):
        window_width = 575
        window_height = 150
        super().__init__(master, width=window_width, height=window_height)
        self.master = master
        self.master.title("GUI")
        self.master.minsize(window_width, window_height)
        self.pack()
        self.target = ProcessingTarget(
            self,
            label_x=30,
            label_y=15,
            bar_x=30,
            bar_y=40,
            bar_len=517,
            )
        self.create_widgets()
        
    def create_widgets(self):
        self.button_start = mgt.tk.ttk.Button(self, text="開始", padding=10, command=self.start_event)
        self.button_start.place(x=235, y=80)
        
    def start_event(self):
        if not self.target.flag_running:
            self.target.start()

def main():
    # button → run via GUI
    window = mgt.tk.Tk()
    app = GuiApplication(master=window)
    app.mainloop()

    # button → direct execution
    #run()

# direct
#main()
