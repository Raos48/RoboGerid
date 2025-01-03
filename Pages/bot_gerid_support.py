#! /usr/bin/env python
#  -*- coding: utf-8 -*-
#
# Support module generated by PAGE version 5.1
#  in conjunction with Tcl version 8.6
#    Oct 06, 2024 09:40:12 PM -03  platform: Windows NT

import sys

import sys
import tkinter as tk
from tkinter import ttk
import bot_gerid_support

def set_Tk_var():
    global combobox
    combobox = tk.StringVar()
    global spinbox
    spinbox = tk.StringVar()

def init(top, gui, *args, **kwargs):
    global w, top_level, root
    w = gui
    top_level = top
    root = top

def destroy_window():
    # Function which closes the window.
    global top_level
    top_level.destroy()
    top_level = None

if __name__ == '__main__':
    import Pages.gerid as gerid
    gerid.vp_start_gui()




