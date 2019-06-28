import re
import pandas as pd


def sub(s):
    s = re.sub(r'\xa0', ' ', s)
    if s == 'nan':
        s = ''
    return s


def read_html(path=''):
    # path = r'data/Demo_2-Fold Dilution/report_resources/abs quant001.html'
    # path = r'data/Demo_Qual Detect Mono Color/report_resources/abs quant001.html'
    with open(path, 'r') as f:
        html = f.read()
    td = pd.read_html(html)[0]
    td = td.applymap(str)
    td = td.values.tolist()
    return td


import os

p = r'E:\part\part_time_job\yg\data\Demo_10-Fold Dilution'
print(os.path.abspath(p))
print(os.path.dirname(p))
print(os.path.split(p))

from tkinter import *
from tkinter.filedialog import askdirectory
import threading
import time

class Gui:

    def __init__(self, root):
        self.root = root
        self.path = StringVar()
        self.file_dir = ''
        Label(self.root, text="File Dir:").grid(row=0, column=0)
        Entry(self.root, textvariable=self.path).grid(row=0, column=1)
        Button(self.root, text="Browse", command=self.selectPath).grid(row=0, column=2)
        self.lb = Label(self.root, text='1')
        self.lb.grid(row=1, column=1)
        Button(self.root, text="Run Extractor", command=self.run).grid(row=2, column=1)

    def selectPath(self):
        self.file_dir = askdirectory()
        self.path.set(self.file_dir)
        print(self.path)
        print(self.file_dir)

    def run(self):
        print(self.file_dir)
        # self.lb.config(text="您选择的文件是：" + self.file_dir)
        self.lb.config(text="run")
        th=threading.Thread(target=self.do)
        th.setDaemon(True)
        th.start()
        # th.join()
        time.sleep(1)
        print('ok')
        # self.lb.config(text='ok')


    def do(self):

        for i in range(10):
            print(i)
            time.sleep(1)


if __name__ == '__main__':
    root = Tk()
    mainui = Gui(root)
    root.mainloop()
