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
import usb.core

# dirección IP y puerto del dispositivo
#ip_address = '192.168.10.179'
ip_address = '192.168.10.66'
#172.16.9.217
port = 1470
# VID = 0X067B
# PID = 0X23A3
# INTERFAZ = ''
# ENDPOINT = ''

Giro_Plato_X = '1e 4d 52 43 58 1f'
Giro_Plato_Y = '1e 4d 52 43 59 1f'

NACK = '10 00 07 00 ff 9f 3a'
ACK = '10 00 07 00 00 81 ca'
RESPUESTA_MARCHA = '10 00 08 00 32 00 FC B0'

'''
def envio_usb(fichero):

    # Conectar al dispositivo USB
    device = usb.core.find(idVendor=VID, idProduct=PID)

    if device is not None:
        try:
            # Abre la interfaz requerida (debes conocer la interfaz de tu dispositivo)
            interface = INTERFAZ  # Reemplaza con el número de interfaz correcto
            endpoint = ENDPOINT   # Reemplaza con el número de endpoint correcto
            device.set_configuration()
            endpoint_address = device[0][(0, 0)][endpoint].bEndpointAddress

            # Envia los bytes al dispositivo
            abrir_fichero_bytes = bytes.fromhex(fichero)
            device.write(endpoint_address, abrir_fichero_bytes, 1000)  # 1000 es el tiempo de espera en milisegundos (puedes ajustarlo según sea necesario)

        finally:
            # Asegúrate de liberar el dispositivo después de su uso
            usb.util.dispose_resources(device)
    else:
        print("Dispositivo no encontrado.")
'''
def reset_marcadora():
    reset_hex = '10 02 07 00 01 7C 83'

    reset_bytes = bytes.fromhex(reset_hex)

    # crea un socket TCP
    client_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)

    # se conecta al dispositivo
    client_socket.connect((ip_address, port))

    # envía el comando al dispositivo
    client_socket.sendall(reset_bytes)

    timeout = 2
    ready = select.select([client_socket], [], [], timeout)

    if ready[0]:

        # espera a recibir la respuesta del dispositivo
        response_bytes = client_socket.recv(1024)
        print('Respuesta recibida/Reiniciando')

        hex_string = ' '.join(['{:02x}'.format(byte) for byte in
                              response_bytes])

        if hex_string == ACK:
            resultado_label2.config(text='Reiniciando Marcadora')
        else:
            resultado_label2.config(text='Error Reset')
    else:
            print('No hay respuesta')
    # cierra la conexión con el dispositivo
    client_socket.close()

def registro_marcado(KH,OF,MARCA):
    try:
        with db_conn_po('ODBC Driver 17 for SQL Server', 'MENITSISTEPLANT', 'marcadongvs', 'landerurgoiti', '2023Urria!') as cursor:
                    # We truncate the table
                    hora=datetime.datetime.now()
                    OFint=int(OF)
                    Registro = f'({KH}, {OFint}, {hora}, {MARCA})'
                    print(Registro)
                    consulta = f'INSERT INTO RegistroMarcado (PartNumber, Orden, FechaMarcado, TipoMarcado) VALUES (?, ?, ?, ?)'
                    cursor.execute(consulta, KH, OFint, hora, MARCA)
                    cursor.commit()
                                    
                    #cursor.execute(consulta, [Registro])
    except po.Error as ex:
        # Capturar la excepción y mostrar información detallada sobre el error
        print("Se produjo un error al insertar los datos en la base de datos:")
        print("Tipo de error:", type(ex).__name__)
        print("Mensaje de error:", ex)
    ventana.focus()
    codigo_textbox.focus_set()


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

def hex_length(hex_text):
    # Elimina los espacios en blanco del texto hexadecimal
    hex_text = re.sub(r'\s', '', hex_text)
    # Calcula la longitud del texto hexadecimal
    length = len(hex_text)//2+2
    LT='{:02X}'.format(length)
    return LT

def calcular_crc(trama_crc):
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

    return crc_ret2

def zero():
    # comando en hexadecimal
    estado_hex = '10 02 0F 00 22 30 30 30 30 30 30 30 30 1C 4F'

    # convierte el comando hexadecimal en bytes
    estado_bytes = bytes.fromhex(estado_hex)

    # crea un socket TCP
    client_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)

    # se conecta al dispositivo
    client_socket.connect((ip_address, port))

    # envía el comando al dispositivo
    client_socket.sendall(estado_bytes)

    # cierra la conexión con el dispositivo
    client_socket.close()

def estado():
    # comando en hexadecimal
    estado_hex = '10 02 07 00 02 4C E0'

    # convierte el comando hexadecimal en bytes
    estado_bytes = bytes.fromhex(estado_hex)

    # crea un socket TCP
    client_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)

    # se conecta al dispositivo
    client_socket.connect((ip_address, port))

    # envía el comando al dispositivo
    client_socket.sendall(estado_bytes)



    # cierra la conexión con el dispositivo
    client_socket.close()


def marcar(OF,NSER,CART,NLOT):
        #BUSCAR LOS DATOS DEL PROGRAMA DEFINIDOS EN LA BASE DE DATOS DE MARCADO
    programa = '91D19600'
    Fichero_hex= ' '.join([hex(ord(c))[2:].zfill(2) for c in programa])

    Carga_Fichero = '10 02 0F 00 22'
    crc_ret2 = calcular_crc(Carga_Fichero + ' ' + Fichero_hex)
    abrir_fichero = Carga_Fichero + ' ' + Fichero_hex + ' ' + crc_ret2
    
    abrir_fichero_bytes = bytes.fromhex(abrir_fichero)
    client_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    client_socket.connect((ip_address, port))
    client_socket.sendall(abrir_fichero_bytes)
    response_bytes = client_socket.recv(1024)
    hex_string = ' '.join(['{:02x}'.format(byte) for byte in response_bytes])

    if hex_string == ACK:
        resultado_label2.config(text='Programa cargado y Marcadora lista')
    elif hex_string == NACK:
        resultado_label2.config(text='Marcadora no disponible')
    else:
        resultado_label2.config(text='ERROR Marcado')

    client_socket.close()

    # cierra la conexión con el dispositivo

    # MARCAR TEXTO

    LT = '0c'

    C1_1_hex = ' '.join([hex(ord(c))[2:].zfill(2) for c in 'F0302    '+CART])
    C2_1_hex = ' '.join([hex(ord(c))[2:].zfill(2) for c in '99UEB    '+NLOT])
    C2_2_hex = ' '.join([hex(ord(c))[2:].zfill(2) for c in '        ' + OF+'-'+NSER])
    
    for i in range(0, 2):

        Marcar_Texto = '10 02 ' + LT + ' 00 36'
    
        '''
        trama = Marcar_Texto + ' ' + Giro_Plato_Y + ' 0a ' \
                + C1_1_hex + ' 0a ' \
                + C2_1_hex + ' 0a ' \
                + C2_2_hex
        '''
        trama = Marcar_Texto + ' ' + C1_1_hex + ' 0a ' \
                + C2_1_hex + ' ' + C2_2_hex
            
        LT = hex_length(trama)
        crc_trama = calcular_crc(trama)

    trama2 = trama + ' ' + crc_trama

    trama_bytes = bytes.fromhex(trama2)
    client_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    client_socket.connect((ip_address, port))
    client_socket.sendall(trama_bytes)
    response_bytes = client_socket.recv(1024)
    hex_string = ' '.join(['{:02x}'.format(byte) for byte in response_bytes])

    time.sleep(5)

    i = 0
    while i < 100 and not hex_string == ACK:
        client_socket.sendall(trama_bytes)
        response_bytes = client_socket.recv(1024)
        hex_string = ' '.join(['{:02x}'.format(byte) for byte in response_bytes]) 
        i=i+1
    if hex_string == ACK:
        resultado_label2.config(text='Texto transmitido, inicio de marcado...')
        print('Marcando...')
        client_socket.close()
        codigo_textbox.insert(0, '')  
        return ACK
    else:
        resultado_label2.config(text='Error en la trama')
        client_socket.close()
        return 0
    


def oficial(OF,NSER):

    with db_conn_po('ODBC Driver 17 for SQL Server', 'MENITBD', 'IDSDMP1', 'lecturaIDS', '0') as cursor:
        try:

            consulta1 = f"SELECT ORDEN.Numero, ORDEN.CodigoArticulo, ORDEN.NLoteCliente, XNSERIE.NumeroSerieCliente \
            FROM ORDEN INNER JOIN XNSERIE ON ORDEN.CodigoArticulo = XNSERIE.CodigoArticulo \
            WHERE (ORDEN.Numero LIKE ?) AND (XNSERIE.NumeroSerieCliente LIKE ?)"
            cursor.execute(consulta1, [OF,NSER])
            resultado = cursor.fetchone()
            CART=str(resultado[1])
            NLOT=str(resultado[2])

        except:
            messagebox.showerror('ERROR', "NO SE ENCUENTRA EL NÚMERO DE SERIE")
            return
    print(OF,NSER,CART,NLOT)
    MAR=marcar(OF,NSER,CART,NLOT)
    time.sleep(5)

    if MAR==ACK:
        hwnd = ventana.winfo_id()
        win32gui.SetForegroundWindow(hwnd)
        codigo_textbox.delete(0, tk.END)
        print('Marcado terminado')
        zero()
        #registro_marcado(KH,OF,'OFICIAL')
        cursor.close()

    else:
        messagebox.showerror('ERROR', "HA HABIDO UN PROBLEMA DURANTE EL MARCADO")


#Crear la ventana
ventana = tk.Tk()
ventana.geometry('450x300+300+300')
ventana.title('PROGRAMA DE MARCADO\nTRANSMISSIONS')

# Botón REPETIR MARCADO
boton_reset = tk.Button(ventana, text='MARCAR', font=('Arial', 10, 'bold'), width=10, height=2, command=lambda: oficial(codigo_textbox.get(),codigo_textbox2.get()))
boton_reset.place(x=120, y=210)
ventana.bind("<Return>", lambda event=None: boton_reset.invoke())

# Botón de cerrar
boton_cerrar = tk.Button(ventana, text='RESET', font=('Arial', 10, 'bold'), width=10, height=2, command=lambda:reset_marcadora())
boton_cerrar.place(x=280, y=210)
ventana.bind("<Return>", lambda event=None: boton_cerrar.invoke())


# label 1
resultado_label1 = tk.Label(ventana, text='  OF  ', font=('Arial', 20, 'bold'), justify='left')
resultado_label1.place(x=60, y=110, anchor=tk.CENTER)

# label 2
resultado_label2 = tk.Label(ventana, text='Nº Serie', font=('Arial', 20, 'bold'), justify='left')
resultado_label2.place(x=60, y=170, anchor=tk.CENTER)

# Crear la caja de texto
codigo_textbox = tk.Entry(ventana, font=('Arial', 20, 'bold'), justify='center')
codigo_textbox.place(x=250, y=110, width=250, height=50, anchor=tk.CENTER)
codigo_textbox.focus_set()
# codigo_textbox.bind('<Return>', lambda event: oficial())

# Crear la caja de texto 2
codigo_textbox2 = tk.Entry(ventana, font=('Arial', 20, 'bold'), justify='center')
codigo_textbox2.place(x=250, y=170, width=250, height=50, anchor=tk.CENTER)
# codigo_textbox.bind('<Return>', lambda event: oficial())

# Cargar la imagen utilizando opencv
imagen1 = Image.open("M:\\OLANET\\G02-13 MITUTOYO CRYSTA APEX\\NGVLEAN\\scripts\\Logo_Egile_Mechanics.jpg")
imagen1 = imagen1.resize((128, 73))
photo = ImageTk.PhotoImage(imagen1)

# label 4
resultado_label_egile = tk.Label(ventana, image=photo)
resultado_label_egile.place(x=70, y=40, anchor=tk.CENTER)

ventana.mainloop()
