import socket
import tkinter as tk
from tkinter import ttk
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
import threading

SISTE = 'MENITSISTEPLANT'
SISTE_BD = 'marcadongvs'
SISTE_USR = 'escrituraCaptor'
SISTE_PWD = '0'
BD = 'MENITBD'
BD_BD = 'IDSDMP1'
BD_USR = 'lecturaIDS'
BD_PWD = '0'


def db_conn_po(driver,server,database,user,password):
    """Generate a database connection with pyodbc
    :param driver: driver to be used
    :param server: server to be connected to
    :param database: database to be connected to
    :param user: user to use for connection
    :param password: password to use for connection
    """
    # We create a connection string with the DRIVER, SERVER, DATABASE, UID AND PWD settings
    conn = po.connect(f"Driver={driver};"
                  f"Server={server};"
                  f"Database={database};"
                  f"UID={user};"
                  f"PWD={password};")
    cursor = conn.cursor()
    return cursor


with db_conn_po('ODBC Driver 17 for SQL Server', SISTE, SISTE_BD, SISTE_USR, SISTE_PWD) as cursor:
    hora=datetime.datetime.now()

    consulta_v = f"SELECT TOP 1 FechaMarcado FROM RegistroMarcado WHERE PartNumber ='VERTICAL' AND TipoMarcado = 'INSPECCIÓN DE MARCADO' ORDER BY FechaMarcado DESC"
    cursor.execute(consulta_v)
    resultado_v = cursor.fetchone()[0]
    consulta_v2 = f"SELECT TOP 10 PartNumber FROM RegistroMarcado ORDER BY FechaMarcado DESC"
    cursor.execute(consulta_v2)
    resultado_v2 = cursor.fetchall()
    lista = len(resultado_v2)
    print(resultado_v2)
    for val in resultado_v2:
        if 'KH37815' in val:
            print('encointrado')
            break
