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
import subprocess
import os
import win32com.client



import glob #para buscar archivos

import os
import win32api
import openpyxl

import tkinter as tk
from tkinter import messagebox
import os
import time
import win32gui
import pyautogui
import shutil


archivo='14144597'
valor='SER '

# Buscar si existe u
# na ventana abierta con el nombre especificado
if not valor.startswith("SER "):
    valor = "SER " + valor

try:
    #hwnd = win32gui.FindWindow(None, 'Inicio de trabajo OF')
    #ventanaIDS = pygetwindow.getWindowsWithTitle('Inicio de trabajo OF')
    hwnd=1
    if hwnd == 0:
        archivo=archivo + '.jpg'
        carpeta=r"\\srv5\Maquinas\Documentacion NGV's\HOJAS DE REGISTRO UNIFICADAS\LOGOS"
        ruta_original = os.path.join(carpeta, archivo)
        ruta_destino = os.path.join(carpeta, "logo.jpg")
        shutil.copy(ruta_original, ruta_destino)
        time.sleep(0.5)

    elif hwnd !=0:
        archivo=archivo + '.jpg'
        carpeta=r"\\srv5\Maquinas\Documentacion NGV's\HOJAS DE REGISTRO UNIFICADAS\LOGOS"
        ruta_original = os.path.join(carpeta, archivo)
        ruta_destino = os.path.join(carpeta, "logo.jpg")
        shutil.copy(ruta_original, ruta_destino)
except:
    messagebox.showerror('ERROR', "ERROR IDS")



