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

# dirección IP y puerto del dispositivo
ip_address = '192.168.10.179'
port = 1470
Data_Matrix_Rect_abre = '1e 44 4d 52 28'
Data_Matrix_Rect_cierra = '29 1f'
DM_SQ_abre = '1e 44 4d 53 28'
DM_SQ_cierra = '29 1f'
GS = '1e 47 53 1f'
RS = '1e 52 53 1f'
Corchete = '5b 29 3e'
EOT = '1e 45 4f 54 1f'
Doce = '31 32'
Punto = '2e'
Activar_Salida_1 = '1e 4d 41 53 31 41 1f'
Activar_Salida_2 = '1e 4d 41 53 32 41 1f'

NACK = '10 00 07 00 ff 9f 3a'
ACK = '10 00 07 00 00 81 ca'
RESPUESTA_MARCHA = '10 00 08 00 32 00 FC B0'

def part_rep():
    
    PNR_window = tk.Toplevel(ventana)
    PNR_window.title("PART-NUMBER REPETIDO")
    PNR_window.geometry('200x150+150+150')
    # Crear los elementos de la ventana de login
    label_partnumber = tk.Label(PNR_window, text="Escanea el Partnumber", font=('Arial',10,'bold'))
    label_partnumber(x=10,y=10)
    entry_partnumber = tk.Entry(PNR_window, font='Arial')
    entry_partnumber.place(x=10,y=30)
    entry_partnumber.bind('<Return>', lambda event: PNR_window.destroy)
    PNR_FUN=entry_partnumber.get()
    PNR_FUN = PNR_FUN.replace('PNR ', '')
    return PNR_FUN


def open_login_window(KH,OF):
    # Crear la ventana de login
    login_window = tk.Toplevel(ventana)
    login_window.title("Login")
    login_window.geometry('200x150+150+150')

    # Crear los elementos de la ventana de login
    label_username = tk.Label(login_window, text="Usuario:", font=('Arial',10,'bold'))
    label_username.place(x=10,y=10)
    entry_username = tk.Entry(login_window, font='Arial')
    entry_username.place(x=10,y=30)
    label_password = tk.Label(login_window, text="Contraseña:", font=('Arial',10,'bold'))
    label_password.place(x=10,y=60)
    entry_password = tk.Entry(login_window, show="*", font='Arial')
    entry_password.place(x=10,y=80)
    button_login = tk.Button(login_window, text="Marcar", command=lambda:(login(entry_username.get(),entry_password.get(),KH,OF),login_window.destroy()))
    login_window.bind("<Return>", lambda event=None: button_login.invoke())
    button_login.place(x=15,y=110)
    login_window.wm_attributes("-topmost", True)

def login(user,pwd,KH,OF):
    username = user
    password = pwd

    if username == "INGENIERIA" and password == "2022URRIA!":
        marcar(KH,OF)
        registro_marcado(KH,OF,'NO OFICIAL')
        codigo_textbox.delete(0, tk.END)
        ventana.focus()
        codigo_textbox.focus_set()
    else:
        messagebox.showerror("Login", "Usuario o contraseña incorrectos")

def registro_marcado(KH,OF,MARCA):
    try:
        with db_conn_po('ODBC Driver 17 for SQL Server', 'MENITSISTEPLANT', 'marcadongvs', 'landerurgoiti', '2023Maiatza!') as cursor:
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



def imprimir_excel(ordenfabri,partnumber):
    try:
        nombre_archivo='Hoja registro ' + partnumber +'.xlsx'
        nombre_impresora='Ineo + 3350 Marcadora NGVS'
        # Busca el archivo en la carpeta actual
        carpeta=r"\\srv5\Maquinas\Documentacion NGV's\HOJAS DE REGISTRO UNIFICADAS"
        ruta_archivo = os.path.join(r"\\srv5\Maquinas\Documentacion NGV's\HOJAS DE REGISTRO UNIFICADAS", nombre_archivo)

        for filename in os.listdir(carpeta):
            if str(partnumber) in filename:
                ruta_archivo = os.path.join(carpeta, filename)

                # Abre el archivo Excel
                libro = openpyxl.load_workbook(ruta_archivo)
                hoja = libro.active

                # Pega el valor en la casilla J1
                hoja.cell(row=1, column=10, value=ordenfabri)

                libro.save(r"\\srv5\Maquinas\Documentacion NGV's\HOJAS DE REGISTRO UNIFICADAS" + 'Temp.xlsx')

                win32api.ShellExecute(
                    0,
                    "print",
                    r"\\srv5\Maquinas\Documentacion NGV's\HOJAS DE REGISTRO UNIFICADAS" + "Temp.xlsx",
                    f'/d:"{nombre_impresora}"',
                    ".",
                    0
                )
                libro.close()
                time.sleep(3)

    except:
        messagebox.showerror('ERROR', "ERROR AL IMPRIMIR LA HOJA DE REGISTRO")

def logo(archivo):
    archivo=archivo + '.jpg'
    
    for i in range(0, 3):

        carpeta=r"\\srv5\Maquinas\Documentacion NGV's\HOJAS DE REGISTRO UNIFICADAS\LOGOS"
        ruta_original = os.path.join(carpeta, archivo)
        ruta_destino = os.path.join(carpeta, "logo.jpg")
        shutil.copy(ruta_original, ruta_destino)
        time.sleep(0.5)

def buscar_y_pegar(valor,archivo):
    # Buscar si existe u
    # na ventana abierta con el nombre especificado
    if not valor.startswith("SER "):
        valor = "SER " + valor
    
    try:
        hwnd = win32gui.FindWindow(None, 'Inicio de trabajo OF')
        #ventanaIDS = pygetwindow.getWindowsWithTitle('Inicio de trabajo OF')
        if hwnd == 0:
            '''
            archivo=archivo + '.jpg'
            carpeta=r"\\srv5\Maquinas\Documentacion NGV's\HOJAS DE REGISTRO UNIFICADAS\LOGOS"
            ruta_original = os.path.join(carpeta, archivo)
            ruta_destino = os.path.join(carpeta, "logo.jpg")
            shutil.copy(ruta_original, ruta_destino)
            time.sleep(0.5)
            '''
            print(archivo)
            logo(archivo)
            shortcut_path = r"C:\Users\taller20\Desktop\PISTOLA.lnk"
            os.startfile(shortcut_path)
            time.sleep(2)
            pyautogui.click(x=100, y=50)
            pyperclip.copy(valor)
            time.sleep(0.2)
            pyautogui.hotkey('ctrl', 'v') # Pegar el valor
            time.sleep(0.2)
            pyautogui.hotkey('enter')
            time.sleep(7)

            return 1
        elif hwnd !=0:
            '''
            archivo=archivo + '.jpg'
            carpeta=r"\\srv5\Maquinas\Documentacion NGV's\HOJAS DE REGISTRO UNIFICADAS\LOGOS"
            ruta_original = os.path.join(carpeta, archivo)
            ruta_destino = os.path.join(carpeta, "logo.jpg")
            shutil.copy(ruta_original, ruta_destino)
            time.sleep(2)
            '''
            print(archivo)
            logo(archivo)
            # Si existe, seleccionar la ventana y pegar el valor
            win32gui.ShowWindow(hwnd, 5)
            win32gui.SetForegroundWindow(hwnd)
            pyperclip.copy(valor)
            time.sleep(0.2)
            pyautogui.hotkey('ctrl', 'v') # Pegar el valor
            time.sleep(0.2)
            pyautogui.hotkey('enter')
            time.sleep(7)

            return 1
    except:
        messagebox.showerror('ERROR', "ERROR IDS")
        return 0

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

def marcar(KH,OF):
        #BUSCAR LOS DATOS DEL PROGRAMA DEFINIDOS EN LA BASE DE DATOS DE MARCADO
    try:
        with db_conn_po('ODBC Driver 17 for SQL Server', 'MENITSISTEPLANT', 'marcadongvs', 'landerurgoiti', '2023Maiatza!') as cursor:
            # We truncate the table
            consulta = 'SELECT PartNumber, Fichero, TipoPieza, Marcadora, TipoMarcado FROM ProgramaMarcado WHERE PartNumber LIKE ?'
            cursor.execute(consulta, [KH])
            resultado = cursor.fetchone() #ESTA FUNCIÓN SACA LA OF, EL CÓDIGO ARTÍCULO Y EL SN DE FUNDIDO
            Fichero=resultado[1]
            ASSY=resultado[2]
            MARCADORA=resultado[3]
            Tipo_Marcado=str(resultado[4])
    except:
        messagebox.showerror('ERROR', "NO HAY UN PROGRAMA DEFINIDO PARA ESTA PIEZA")
        return

    Fichero_hex= ' '.join([hex(ord(c))[2:].zfill(2) for c in Fichero])

    Carga_Fichero = '10 02 0F 00 22'
    crc_ret2 = calcular_crc(Carga_Fichero + ' ' + Fichero_hex)
    abrir_fichero = Carga_Fichero + ' ' + Fichero_hex + ' ' + crc_ret2
    #
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
    SER = 'SER 201293DMP' + OF
    SER_hex = ' '.join([hex(ord(c))[2:].zfill(2) for c in SER])
    MFR = 'MFR K0680'
    MFR_hex = ' '.join([hex(ord(c))[2:].zfill(2) for c in MFR])
    PNR = 'PNR ' + KH
    PNR_hex = ' '.join([hex(ord(c))[2:].zfill(2) for c in PNR])
    if ASSY=='ASSY':
        ASSY_hex = ' '.join([hex(ord(c))[2:].zfill(2) for c in ASSY])
    else:
        ASSY_hex='20'
    
    DM = Corchete + ' ' + RS + ' ' + Doce + ' ' + GS + ' ' + MFR_hex + ' ' + GS + ' ' + SER_hex + ' ' + GS + ' ' + PNR_hex + ' ' + RS + ' ' + EOT

    if KH=='KH37831' or KH=='KH37833':
        SER = '201293DMP' + OF
        SER_hex2 = ' '.join([hex(ord(c))[2:].zfill(2) for c in SER])
        PNR = KH
        PNR_hex2 = ' '.join([hex(ord(c))[2:].zfill(2) for c in PNR])
    else:
        SER_hex2=SER_hex
        PNR_hex2=PNR_hex
    for i in range(0, 2):

        if MARCADORA=='VERTICAL':
            Marcar_Texto = '10 02 ' + LT + ' 00 36'

        elif MARCADORA=='HORIZONTAL':
            Marcar_Texto = '10 02 ' + LT + ' 00 36' + ' ' + Activar_Salida_1
        
        if Tipo_Marcado == '1':#tipo de marcado con DM único y rectangular

            trama = Marcar_Texto + ' ' + Punto + ' 0a ' \
                    + Data_Matrix_Rect_abre + ' ' + DM + ' ' + Data_Matrix_Rect_cierra + ' 0a ' \
                    + MFR_hex + ' 0a ' \
                    + SER_hex + ' 0a ' \
                    + '0a ' \
                    + PNR_hex + ' 0a ' \
                    + ASSY_hex

        elif Tipo_Marcado == '2':#tipo de marcado con DM único y cradrado

            trama = Marcar_Texto + ' ' + Punto + ' 0a ' \
                    + DM_SQ_abre + ' ' + DM + ' ' + DM_SQ_cierra + ' 0a ' \
                    + MFR_hex + ' 0a ' \
                    + SER_hex + ' 0a ' \
                    + '0a ' \
                    + PNR_hex + ' 0a ' \
                    + ASSY_hex
            
        elif Tipo_Marcado=='3':#tipo de marcado antiguo, con 2 DM cradrados
            DM1 = MFR_hex + ' 20 ' + SER_hex
            DM2 = PNR_hex + ' 20 ' + ASSY_hex
            trama = Marcar_Texto + ' ' + Punto + ' 0a ' \
                    + DM_SQ_abre + ' ' + DM1 + ' ' + DM_SQ_cierra + ' 0a ' \
                    + MFR_hex + ' 0a ' \
                    + SER_hex2 + ' 0a ' \
                    + DM_SQ_abre + ' ' + DM2 + ' ' + DM_SQ_cierra + ' 0a ' \
                    + PNR_hex2 + ' 0a ' \
                    + ASSY_hex
            
        if MARCADORA=='HORIZONTAL':
            #AÑADIR EL PUNTO DE SALIDA PARA QUE NO GOLPEE CON EL NGV
            trama = trama + ' 0a ' \
            + Punto

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
    
def no_oficial(coditex):

    SER_FUN = coditex.replace('SER', '')
    SER_FUN = SER_FUN.replace(' ', '')

    with db_conn_po('ODBC Driver 17 for SQL Server', 'MENITBD', 'IDSDMP1', 'lecturaIDS', '0') as cursor:
        try:
            # COMPROBAR SI EL NUMERO DE SERIE DE FUNDIDO ESTÁ DUPLICADO DENTRO DE LAS OFS RESERVADAS, ABIERTAS O CERRADAS 
            consulta0 = f"SELECT COUNT(*) FROM PARTMAT INNER JOIN ORDEN ON PARTMAT.NumeroOrden = ORDEN.Numero WHERE (PARTMAT.NSerieMaterial LIKE ?)"
            #consulta0 = 'SELECT COUNT(*) FROM PARTMAT WHERE NSerieMaterial LIKE ?'
            cursor.execute(consulta0, [SER_FUN])
            resultado = cursor.fetchone()[0]

            if resultado==1:
                #SI SOLO HAY UNA, EXTRAER OF, CÓDIGO ARTÍCULO Y PARTNUMBER 
                # EN ESTE CASO SE PUEDEN MARCAR PIEZAS ENTRE LAS RESERVADAS Y LAS ABIERTAS
                consulta1 = f"SELECT PARTMAT.NumeroOrden, PARTMAT.CodigoComponente FROM PARTMAT INNER JOIN ORDEN ON PARTMAT.NumeroOrden = ORDEN.Numero WHERE (PARTMAT.NSerieMaterial LIKE ?)"
                cursor.execute(consulta1, [SER_FUN])
                resultado = cursor.fetchone()
                OF=str(resultado[0])
                Codi_Art=str(resultado[1])
                consulta2='SELECT CodigoAlternativo FROM ARTICULO WHERE CODIGO LIKE ?'
                cursor.execute(consulta2, [Codi_Art])
                resultado2 = cursor.fetchone()
                KH = resultado2[0]
                        
            elif resultado >= 2:
                PN_FUN=part_rep()

                consulta3=f"SELECT Codigo FROM ARTICULO WHERE CodigoAlternativo LIKE ?"
                cursor.execute(consulta3, [PN_FUN])
                resultado3 = cursor.fetchone()
                CO_MA = resultado3[0]

                consulta1 = f"SELECT PARTMAT.NumeroOrden, PARTMAT.CodigoComponente FROM PARTMAT INNER JOIN ORDEN ON PARTMAT.NumeroOrden = ORDEN.Numero WHERE (PARTMAT.CodigoMaterial={CO_MA}) AND (PARTMAT.NSerieMaterial LIKE ?)"
                cursor.execute(consulta1, [SER_FUN])
                resultado = cursor.fetchone()
                OF=str(resultado[0])
                Codi_Art=str(resultado[1])
                consulta2='SELECT CodigoAlternativo FROM ARTICULO WHERE CODIGO LIKE ?'
                cursor.execute(consulta2, [Codi_Art])
                resultado2 = cursor.fetchone()
                KH = resultado2[0]

        except:
            messagebox.showerror('ERROR', "NO SE ENCUENTRA EL NÚMERO DE SERIE \n OF NO RESERVADA")
            return
    
    open_login_window(KH,OF)
    estado()
    cursor.close()  
    ventana.focus()
    codigo_textbox.focus_set()

def soloHOJA(coditex):

    SER_FUN = coditex.replace('SER', '')
    SER_FUN = SER_FUN.replace(' ', '')

    with db_conn_po('ODBC Driver 17 for SQL Server', 'MENITBD', 'IDSDMP1', 'lecturaIDS', '0') as cursor:
        try:
            # COMPROBAR SI EL NUMERO DE SERIE DE FUNDIDO ESTÁ DUPLICADO DENTRO DE LAS OFS RESERVADAS, ABIERTAS O CERRADAS 
            consulta0 = f"SELECT COUNT(*) FROM PARTMAT INNER JOIN ORDEN ON PARTMAT.NumeroOrden = ORDEN.Numero WHERE (PARTMAT.NSerieMaterial LIKE ?)"
            #consulta0 = 'SELECT COUNT(*) FROM PARTMAT WHERE NSerieMaterial LIKE ?'
            cursor.execute(consulta0, [SER_FUN])
            resultado = cursor.fetchone()[0]

            if resultado==1:
                #SI SOLO HAY UNA, EXTRAER OF, CÓDIGO ARTÍCULO Y PARTNUMBER 
                # EN ESTE CASO SE PUEDEN MARCAR PIEZAS ENTRE LAS RESERVADAS Y LAS ABIERTAS
                consulta1 = f"SELECT PARTMAT.NumeroOrden, PARTMAT.CodigoComponente FROM PARTMAT INNER JOIN ORDEN ON PARTMAT.NumeroOrden = ORDEN.Numero WHERE (PARTMAT.NSerieMaterial LIKE ?)"
                cursor.execute(consulta1, [SER_FUN])
                resultado = cursor.fetchone()
                OF=str(resultado[0])
                Codi_Art=str(resultado[1])
                consulta2='SELECT CodigoAlternativo FROM ARTICULO WHERE CODIGO LIKE ?'
                cursor.execute(consulta2, [Codi_Art])
                resultado2 = cursor.fetchone()
                KH = resultado2[0]
                        
            elif resultado >= 2:
                PN_FUN=part_rep()

                consulta3=f"SELECT Codigo FROM ARTICULO WHERE CodigoAlternativo LIKE ?"
                cursor.execute(consulta3, [PN_FUN])
                resultado3 = cursor.fetchone()
                CO_MA = resultado3[0]

                consulta1 = f"SELECT PARTMAT.NumeroOrden, PARTMAT.CodigoComponente FROM PARTMAT INNER JOIN ORDEN ON PARTMAT.NumeroOrden = ORDEN.Numero WHERE (PARTMAT.CodigoMaterial={CO_MA}) AND (PARTMAT.NSerieMaterial LIKE ?)"
                cursor.execute(consulta1, [SER_FUN])
                resultado = cursor.fetchone()
                OF=str(resultado[0])
                Codi_Art=str(resultado[1])
                consulta2='SELECT CodigoAlternativo FROM ARTICULO WHERE CODIGO LIKE ?'
                cursor.execute(consulta2, [Codi_Art])
                resultado2 = cursor.fetchone()
                KH = resultado2[0]

        except:
            messagebox.showerror('ERROR', "NO SE ENCUENTRA EL NÚMERO DE SERIE \n OF NO RESERVADA")
            return

    imprimir_excel(OF,KH)

def soloOF(coditex):

    SER_FUN = coditex.replace('SER', '')
    SER_FUN = SER_FUN.replace(' ', '')

    with db_conn_po('ODBC Driver 17 for SQL Server', 'MENITBD', 'IDSDMP1', 'lecturaIDS', '0') as cursor:
        try:
            # COMPROBAR SI EL NUMERO DE SERIE DE FUNDIDO ESTÁ DUPLICADO DENTRO DE LAS OFS RESERVADAS, ABIERTAS O CERRADAS 
            consulta0 = f"SELECT COUNT(*) FROM PARTMAT INNER JOIN ORDEN ON PARTMAT.NumeroOrden = ORDEN.Numero WHERE (PARTMAT.NSerieMaterial LIKE ?)"
            #consulta0 = 'SELECT COUNT(*) FROM PARTMAT WHERE NSerieMaterial LIKE ?'
            cursor.execute(consulta0, [SER_FUN])
            resultado = cursor.fetchone()[0]

            if resultado==1:
                #SI SOLO HAY UNA, EXTRAER OF, CÓDIGO ARTÍCULO Y PARTNUMBER 
                # EN ESTE CASO SE PUEDEN MARCAR PIEZAS ENTRE LAS RESERVADAS Y LAS ABIERTAS
                consulta1 = f"SELECT PARTMAT.NumeroOrden, PARTMAT.CodigoComponente FROM PARTMAT INNER JOIN ORDEN ON PARTMAT.NumeroOrden = ORDEN.Numero WHERE (PARTMAT.NSerieMaterial LIKE ?)"
                cursor.execute(consulta1, [SER_FUN])
                resultado = cursor.fetchone()
                OF=str(resultado[0])
                Codi_Art=str(resultado[1])
                consulta2='SELECT CodigoAlternativo FROM ARTICULO WHERE CODIGO LIKE ?'
                cursor.execute(consulta2, [Codi_Art])
                resultado2 = cursor.fetchone()
                KH = resultado2[0]
                        
            elif resultado >= 2:
                PN_FUN=part_rep()

                consulta3=f"SELECT Codigo FROM ARTICULO WHERE CodigoAlternativo LIKE ?"
                cursor.execute(consulta3, [PN_FUN])
                resultado3 = cursor.fetchone()
                CO_MA = resultado3[0]

                consulta1 = f"SELECT PARTMAT.NumeroOrden, PARTMAT.CodigoComponente FROM PARTMAT INNER JOIN ORDEN ON PARTMAT.NumeroOrden = ORDEN.Numero WHERE (PARTMAT.CodigoMaterial={CO_MA}) AND (PARTMAT.NSerieMaterial LIKE ?)"
                cursor.execute(consulta1, [SER_FUN])
                resultado = cursor.fetchone()
                OF=str(resultado[0])
                Codi_Art=str(resultado[1])
                consulta2='SELECT CodigoAlternativo FROM ARTICULO WHERE CODIGO LIKE ?'
                cursor.execute(consulta2, [Codi_Art])
                resultado2 = cursor.fetchone()
                KH = resultado2[0]

        except:
            messagebox.showerror('ERROR', "NO SE ENCUENTRA EL NÚMERO DE SERIE \n OF NO RESERVADA")
            return

    IDS=buscar_y_pegar('SER ' + SER_FUN,Codi_Art)#Abre la OF y la Imprime, necesita el SER y el codigo de articulo para el logo de color
    if IDS==0:
        return


def oficial(coditex):

    SER_FUN = coditex.replace('SER', '')
    SER_FUN = SER_FUN.replace(' ', '')

    with db_conn_po('ODBC Driver 17 for SQL Server', 'MENITBD', 'IDSDMP1', 'lecturaIDS', '0') as cursor:
        try:
            # COMPROBAR SI EL NUMERO DE SERIE DE FUNDIDO ESTÁ DUPLICADO DENTRO DE LAS OFS RESERVADAS
            consulta0 = f"SELECT COUNT(*) FROM PARTMAT INNER JOIN ORDEN ON PARTMAT.NumeroOrden = ORDEN.Numero WHERE (PARTMAT.NSerieMaterial LIKE ?) AND (ORDEN.Situacion = 'Reservada')"
            #consulta0 = 'SELECT COUNT(*) FROM PARTMAT WHERE NSerieMaterial LIKE ?'
            cursor.execute(consulta0, [SER_FUN])
            resultado = cursor.fetchone()[0]

            if resultado==1:
                #SI SOLO HAY UNA, EXTRAER OF, CÓDIGO ARTÍCULO Y PARTNUMBER DENTRO DE LAS RESERVADAS
                consulta1 = f"SELECT PARTMAT.NumeroOrden, PARTMAT.CodigoComponente FROM PARTMAT INNER JOIN ORDEN ON PARTMAT.NumeroOrden = ORDEN.Numero WHERE (PARTMAT.NSerieMaterial LIKE ?) AND (ORDEN.Situacion = 'Reservada')"
                cursor.execute(consulta1, [SER_FUN])
                resultado = cursor.fetchone()
                OF=str(resultado[0])
                Codi_Art=str(resultado[1])
                consulta2=f"SELECT CodigoAlternativo FROM ARTICULO WHERE CODIGO LIKE ?"
                cursor.execute(consulta2, [Codi_Art])
                resultado2 = cursor.fetchone()
                KH = resultado2[0]
                        
            elif resultado >= 2:
                PN_FUN=part_rep()

                consulta3=f"SELECT Codigo FROM ARTICULO WHERE CodigoAlternativo LIKE ?"
                cursor.execute(consulta3, [PN_FUN])
                resultado3 = cursor.fetchone()
                CO_MA = resultado3[0]

                consulta1 = f"SELECT PARTMAT.NumeroOrden, PARTMAT.CodigoComponente FROM PARTMAT INNER JOIN ORDEN ON PARTMAT.NumeroOrden = ORDEN.Numero WHERE (PARTMAT.CodigoMaterial={CO_MA}) AND (PARTMAT.NSerieMaterial LIKE ?)"
                cursor.execute(consulta1, [SER_FUN])
                resultado = cursor.fetchone()
                OF=str(resultado[0])
                Codi_Art=str(resultado[1])
                consulta2='SELECT CodigoAlternativo FROM ARTICULO WHERE CODIGO LIKE ?'
                cursor.execute(consulta2, [Codi_Art])
                resultado2 = cursor.fetchone()
                KH = resultado2[0]
            elif resultado==0:
                messagebox.showerror('ERROR', "OF NO RESERVADA")
                return

        except:
            messagebox.showerror('ERROR', "NO SE ENCUENTRA EL NÚMERO DE SERIE")
            return
    
    #COMPROBAR SI HAY REGISTRO SELECT EXISTS(SELECT * FROM nombre_de_la_tabla WHERE nombre_de_la_columna = valor_buscado);
    #try:
    with db_conn_po('ODBC Driver 17 for SQL Server', 'MENITSISTEPLANT', 'marcadongvs', 'landerurgoiti', '2023Maiatza!') as cursor:
        consulta = 'SELECT COUNT(*) FROM RegistroMarcado WHERE Orden LIKE ?'
        cursor.execute(consulta, [OF])
        resultado = cursor.fetchone()[0]
    if resultado!=0:
        messagebox.showerror('ERROR', "ESTA PIEZA YA SE HA MARCADO")
        return
    #except:
        #messagebox.showerror('ERROR', "ERROR EN LA COMPROBACIÓN DEL REGISTRO \n AVISAR A INGENIERÍA")
        #return
    #ABRIR LA OF EN IDS Y IMPRIMIR LA HOJA DE LA OF
    
    IDS=buscar_y_pegar('SER ' + SER_FUN,Codi_Art)#Abre la OF y la Imprime, necesita el SER y el codigo de articulo para el logo de color
    if IDS==0:
        return
    
    #IMPRIMIR HOJA DE REGISTRO

    imprimir_excel(OF,KH)

    MAR=marcar(KH,OF)
    time.sleep(40)
    if MAR==ACK:
        hwnd = ventana.winfo_id()
        win32gui.SetForegroundWindow(hwnd)
        codigo_textbox.delete(0, tk.END)
        resultado_label2.config(text='Marcado terminado')
        zero()
        registro_marcado(KH,OF,'OFICIAL')
        cursor.close()

    else:
        messagebox.showerror('ERROR', "HA HABIDO UN PROBLEMA DURANTE EL MARCADO")


#Crear la ventana
ventana = tk.Tk()
ventana.geometry('450x300+300+300')
ventana.title('PROGRAMA DE MARCADO')

# Botón REPETIR MARCADO
boton_reset = tk.Button(ventana, text='REPETIR\nMARCADO', width=10, height=2, command=lambda: no_oficial(codigo_textbox.get()))
boton_reset.place(x=100, y=155)
ventana.bind("<Return>", lambda event=None: boton_reset.invoke())

# Botón de cerrar
boton_cerrar = tk.Button(ventana, text='CERRAR', width=10, height=2, command=ventana.destroy)
boton_cerrar.place(x=200, y=155)

# Botón SOLO OF
boton_reset = tk.Button(ventana, text='OF', width=12, height=1, command=lambda: soloOF(codigo_textbox.get()))
boton_reset.place(x=300, y=150)

# Botón SOLO HOJA DE REGISTRO
boton_reset = tk.Button(ventana, text='HOJA REGISTRO', width=12, height=1, command=lambda: soloHOJA(codigo_textbox.get()))
boton_reset.place(x=300, y=180)

# label 1
resultado_label1 = tk.Label(ventana, text='SER ->', font=('Arial', 20, 'bold'), justify='left')
resultado_label1.place(x=80, y=110, anchor=tk.CENTER)

# Crear la caja de texto
codigo_textbox = tk.Entry(ventana, font=('Arial', 20, 'bold'), justify='center')
codigo_textbox.place(x=250, y=110, width=250, height=50, anchor=tk.CENTER)
codigo_textbox.focus_set()
codigo_textbox.bind('<Return>', lambda event: oficial(codigo_textbox.get()))

# label 2
resultado_label2 = tk.Label(ventana, text='', font=('Arial', 10), justify='left')
resultado_label2.place(x=250, y=220, anchor=tk.CENTER)

# Cargar la imagen utilizando opencv
imagen1 = Image.open("M:\\OLANET\\G02-13 MITUTOYO CRYSTA APEX\\NGVLEAN\\scripts\\Logo_Egile_Mechanics.jpg")
imagen1 = imagen1.resize((128, 73))
photo = ImageTk.PhotoImage(imagen1)

# label 4
resultado_label_egile = tk.Label(ventana, image=photo)
resultado_label_egile.place(x=70, y=40, anchor=tk.CENTER)

estado()

ventana.mainloop()
