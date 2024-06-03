import threading
import socket
import tkinter as tk
import tkinter.messagebox as messagebox
from PIL import ImageTk, Image
import time
import re
import select
import pyodbc as po
import pyautogui
import pygetwindow
import pyperclip
import datetime
import os
import win32api
import win32gui
import openpyxl
import shutil


hwnd = win32gui.FindWindow(None, 'Gestor de Aplicaciones Comerciales e Industriales')
#ventanaIDS = pygetwindow.getWindowsWithTitle('Inicio de trabajo OF')
if hwnd == 0:
    #shortcut_path = r"C:\Users\taller20\Desktop\IDSWIN.lnk"
    shortcut_path = r"C:\Users\lander.urgoiti\Desktop\IDSWIN.lnk"
    os.startfile(shortcut_path)
    time.sleep(3)
    pyautogui.click(x=800, y=450)
    time.sleep(7)
    pyautogui.click(x=150, y=40)
    time.sleep(0.2)
    pyautogui.click(x=150, y=270)                      
    time.sleep(0.2)
    pyautogui.click(x=900, y=340)
    time.sleep(0.2)   
    pyautogui.click(x=900, y=380)          
    #time.sleep(0.5)     
    #pyautogui.click(x=900, y=450)
    pyperclip.copy('333333')
    pyautogui.hotkey('tab')
    pyautogui.hotkey('tab')
    pyautogui.hotkey('tab')
    pyautogui.hotkey('ctrl', 'v')
    pyautogui.hotkey('tab')
    pyperclip.copy('333333')
    pyautogui.hotkey('ctrl', 'v')
    pyautogui.hotkey('tab')
    pyautogui.hotkey('tab')
    pyautogui.hotkey('tab')
    pyautogui.hotkey('tab')
    pyautogui.hotkey('tab')
    pyautogui.hotkey('tab')
    pyautogui.hotkey('tab')
    pyautogui.hotkey('tab')
    pyautogui.hotkey('tab')
    pyautogui.hotkey('tab')
    pyautogui.hotkey('enter')
    time.sleep(4)
    pyautogui.click(x=500, y=50)
    pyautogui.hotkey('tab')
    pyautogui.hotkey('tab')
    pyautogui.hotkey('tab')
    pyautogui.hotkey('tab')
    pyautogui.hotkey('tab')
    pyautogui.hotkey('enter')
    time.sleep(5)