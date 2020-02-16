#! /usr/bin/env python
#  -*- coding: utf-8 -*-
#
# GUI module generated by PAGE version 4.22
#  in conjunction with Tcl version 8.6
#    May 29, 2019 11:43:32 PM -04  platform: Windows NT

import sys

try:
    import Tkinter as tk
except ImportError:
    import tkinter as tk

try:
    import ttk
    py3 = False
except ImportError:
    import tkinter.ttk as ttk
    py3 = True

import genera_png_support

def vp_start_gui():
    '''Starting point when module is the main routine.'''
    global val, w, root
    root = tk.Tk()
    top = Toplevel1 (root)
    genera_png_support.init(root, top)
    root.mainloop()

w = None
def create_Toplevel1(root, *args, **kwargs):
    '''Starting point when module is imported by another program.'''
    global w, w_win, rt
    rt = root
    w = tk.Toplevel (root)
    top = Toplevel1 (w)
    genera_png_support.init(w, top, *args, **kwargs)
    return (w, top)

def destroy_Toplevel1():
    global w
    w.destroy()
    w = None

class Toplevel1:
    def __init__(self, top=None):
        '''This class configures and populates the toplevel window.
           top is the toplevel containing window.'''
        _bgcolor = '#d9d9d9'  # X11 color: 'gray85'
        _fgcolor = '#000000'  # X11 color: 'black'
        _compcolor = '#d9d9d9' # X11 color: 'gray85'
        _ana1color = '#d9d9d9' # X11 color: 'gray85'
        _ana2color = '#ececec' # Closest X11 color: 'gray92'

        top.geometry("435x389+650+150")
        top.title("New Toplevel")
        top.configure(background="#d9d9d9")

        self.Listbox1 = tk.Listbox(top)
        self.Listbox1.place(relx=0.046, rely=0.051, relheight=0.648
                , relwidth=0.906)
        self.Listbox1.configure(background="white")
        self.Listbox1.configure(disabledforeground="#a3a3a3")
        self.Listbox1.configure(font="-family {Courier New} -size 10")
        self.Listbox1.configure(foreground="#000000")
        self.Listbox1.configure(width=394)

        self.Label1 = tk.Label(top)
        self.Label1.place(relx=0.046, rely=0.72, height=21, width=394)
        self.Label1.configure(background="#d9d9d9")
        self.Label1.configure(disabledforeground="#a3a3a3")
        self.Label1.configure(foreground="#000000")
        self.Label1.configure(text='''%d recibos seleccionados''')
        self.Label1.configure(width=394)

        self.Entry1 = tk.Entry(top)
        self.Entry1.place(relx=0.368, rely=0.797,height=20, relwidth=0.331)
        self.Entry1.configure(background="white")
        self.Entry1.configure(disabledforeground="#a3a3a3")
        self.Entry1.configure(font="-family {Courier New} -size 10")
        self.Entry1.configure(foreground="#000000")
        self.Entry1.configure(insertbackground="black")
        self.Entry1.configure(width=144)

        self.Label2 = tk.Label(top)
        self.Label2.place(relx=0.069, rely=0.797, height=21, width=124)
        self.Label2.configure(anchor='w')
        self.Label2.configure(background="#d9d9d9")
        self.Label2.configure(disabledforeground="#a3a3a3")
        self.Label2.configure(foreground="#000000")
        self.Label2.configure(text='''Recibos posteriores al:''')
        self.Label2.configure(width=124)

        self.btAcepta = tk.Button(top)
        self.btAcepta.place(relx=0.667, rely=0.9, height=24, width=48)
        self.btAcepta.configure(activebackground="#ececec")
        self.btAcepta.configure(activeforeground="#000000")
        self.btAcepta.configure(background="#d9d9d9")
        self.btAcepta.configure(disabledforeground="#a3a3a3")
        self.btAcepta.configure(foreground="#000000")
        self.btAcepta.configure(highlightbackground="#d9d9d9")
        self.btAcepta.configure(highlightcolor="black")
        self.btAcepta.configure(pady="0")
        self.btAcepta.configure(text='''Acepta''')

        self.btCancela = tk.Button(top)
        self.btCancela.place(relx=0.805, rely=0.9, height=24, width=53)
        self.btCancela.configure(activebackground="#ececec")
        self.btCancela.configure(activeforeground="#000000")
        self.btCancela.configure(background="#d9d9d9")
        self.btCancela.configure(disabledforeground="#a3a3a3")
        self.btCancela.configure(foreground="#000000")
        self.btCancela.configure(highlightbackground="#d9d9d9")
        self.btCancela.configure(highlightcolor="black")
        self.btCancela.configure(pady="0")
        self.btCancela.configure(text='''Cancela''')

if __name__ == '__main__':
    vp_start_gui()





