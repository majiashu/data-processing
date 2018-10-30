# -*- encoding: utf-8 -*-
"""
testGUI.py
Created on 2018/8/24 14:22
Copyright (c) 2018/8/24, 
@author: 马家树(majstx@163.com)
"""
from test.cleanexcelrow import *
import tkinter as tk

window = tk.Tk()
window.title('FormatFile')
# 窗口尺寸
window.geometry('300x200')
# 显示出来

e1 = tk.Entry(window, show=None)
e1.pack()

e2 = tk.Entry(window, show=None)
e2.pack()


def get_path():
    varin = e1.get()
    valout = e2.get()
    print(varin, valout)
    main(varin, valout)


b1 = tk.Button(window, text="确定", width=10, height=1, command=get_path)
b1.pack()
b2 = tk.Button(window, text="退出", width=10, height=1, command=window.quit)
b2.pack()

window.mainloop()
