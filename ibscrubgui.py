import tkinter
from tkinter import ttk
from tkinter.filedialog import askdirectory
import os
import os.path
import time
import ib
import sys
import threading


class StdoutRedirector(object):
    def __init__(self,text_widget):
        self.text_space = text_widget

    def write(self,string):
        self.text_space.insert('end', string)
        self.text_space.see('end')


class IbScrubGui(tkinter.Tk):
    def __init__(self, parent):
        tkinter.Tk.__init__(self, parent)
        self.parent = parent
        self.dir_opt = options = {}
        self.directory = ''
        options['initialdir'] = os.path.join(os.path.expanduser('~'), "Documents")
        options['mustexist'] = True
        options['parent'] = parent
        options['title'] = 'This is a title'

        self.working_dir_var = tkinter.StringVar()
        self.working_dir_label_var = tkinter.StringVar()
        self.select_button = ttk.Button(self, text="Choose Source Directory",
                                       command=self.on_select_button_click)
        self.select_button.grid(column=0, row=0, sticky='NW', padx=10, pady=10)

        self.generate_button = ttk.Button(self, text="Generate IB Scrub Report",
                                       command=self.on_generate_button_click)
        self.generate_button.grid(column=1, row=0, sticky='NE', padx=10, pady=10)

        self.text_box = tkinter.Text(self.parent, wrap='word', height=28, width=50)
        self.text_box.grid(column=0, row=2, columnspan=2, sticky='NSWE', padx=5, pady=5)
        sys.stderr = StdoutRedirector(self.text_box)
        self.active_thread = None
        # self.include_flash = tkinter.IntVar()
        # self.new_asups = tkinter.IntVar()

        self.initialize()

    def initialize(self):
        self.grid()

        self.working_dir_var.set("Select working directory")
        # label = tkinter.Label(self, textvariable=self.working_dir_label_var, anchor="w", fg="white", bg="blue")

        style = ttk.Style()
        style.configure("BW.TLabel", anchor="w", foreground="white", background="blue")
        label = ttk.Label(style="BW.TLabel", textvariable=self.working_dir_label_var)

        label.grid(column=0, row=1, columnspan=2, sticky='EW', padx=10, pady=10)
        self.working_dir_label_var.set("Please select working directory with IB Details, Service Contracts Expiring, score.xlsx, etc")
        # ttk.Checkbutton(self, text="Include Flash Tab", variable=self.include_flash).grid(row=3, sticky='W')
        # ttk.Checkbutton(self, text="Get New Asups", variable=self.new_asups).grid(row=4, sticky='W')

        self.generate_button['state'] = 'disabled'
        self.grid_columnconfigure(0, weight=1)
        self.resizable(True, False)
        self.minsize(width=1000, height=600)
        self.update()
        self.geometry(self.geometry())       

    def on_select_button_click(self):
        if not self.active_thread or not self.active_thread.is_alive():
            self.directory = askdirectory()
            if self.directory:
                os.chdir(self.directory)
                self.working_dir_label_var.set(self.directory)
                # with open("thisisatest.txt", 'w') as f:
                #     print(self.directory, file=f)
                # print("You clicked the select button")
                self.generate_button['state'] = 'normal'

    def on_generate_button_click(self):
        if not self.active_thread or not self.active_thread.is_alive():
            # self.generate_button['state'] = 'disabled'
            # self.select_button['state'] = 'disabled'
            # ib_report = ib.IbDetails()
            print("Starting...", file=sys.stderr)
            # if self.include_flash.get():
            #     include_flash = True
            # else:
            #     include_flash = False
            # if self.new_asups.get():
            #     new_asups = True
            # else:
            #     new_asups = False

            self.active_thread = threading.Thread(target=ib.IbDetails)
            self.active_thread.start()
            # time.sleep(1)
            # os.system("start " + os.path.join(self.directory, "thisisatest.txt"))
            # os.system("start " + os.path.join(self.directory, ib_report.get_ib_report_name()))
            # print("You clicked the generate button")
            # self.generate_button['state'] = 'normal'
            # self.select_button['state'] = 'normal'

if __name__ == "__main__":
    app = IbScrubGui(None)
    app.title('IB Team Installed Base Scrub v1.6.7')
    app.mainloop()
