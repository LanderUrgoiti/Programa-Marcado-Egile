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
import shutil

trama_crc='10 02 0F 00 22 30 30 30 30 30 30 30 30'

trama_bytes = bytes.fromhex(trama_crc)  # Convertir la trama hexadecimal en una cadena de bytes
crc = 0x0000
for byte in trama_bytes:
    crc ^= byte << 8  # Hacer XOR del byte con el valor del CRC desplazado 8 bits hacia la izquierda
    for _ in range(8):
        if crc & 0x8000:  # Comprobar si el bit más significativo está encendido
            crc = (crc << 1) ^ 0x1021  # Aplicar el polinomio si el bit más significativo está encendido
        else:
            crc <<= 1  # Si el bit más significativo está apagado, simplemente desplazar a la izquierda
    crc &= 0xFFFF  # Asegurarse de que el valor del CRC sigue siendo de 16 bits

crc_ret = '{:04x}'.format(crc)  # Devolver el valor del CRC en formato de cadena sin el prefijo "0x"
crc_ret2 = ' '.join(crc_ret[i:i + 2] for i in range(0, len(crc_ret), 2))
print(crc_ret,crc_ret2)