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

# dirección IP y puerto del dispositivo
ip_address = '192.168.10.179'
port = 1470
Data_Matrix_Rect_abre = '1e 44 4d 52 28'
Data_Matrix_Rect_cierra = '29 1f'
DM_SQ_abre = '1e 44 4d 53 28'
DM_SQ_cierra = '29 1f'
GS = '1e 47 53 1f'
GS_regi = '<0x1d>'
RS = '1e 52 53 1f'
RS_regi = '<0x1e>'
Corchete = '5b 29 3e'
Corchete_regi = '[)>'
EOT = '1e 45 4f 54 1f'
EOT_regi = '<0x04>'
Doce = '31 32'
Punto = '2e'
Activar_Salida_1 = '1e 4d 41 53 31 41 1f'
Activar_Salida_2 = '1e 4d 41 53 32 41 1f'

NACK = '10 00 07 00 ff 9f 3a'
ACK = '10 00 07 00 00 81 ca'
RESPUESTA_MARCHA = '10 00 08 00 32 00 FC B0'

SISTE = 'MENITSISTEPLANT'
SISTE_BD = 'marcadongvs'
SISTE_USR = 'escrituraCaptor'
SISTE_PWD = '0'
BD = 'MENITBD'
BD_BD = 'IDSDMP1'
BD_USR = 'lecturaIDS'
BD_PWD = '0'

def control(vert, hori):

    while True:
        try:
            with db_conn_po('ODBC Driver 17 for SQL Server', SISTE, SISTE_BD, SISTE_USR, SISTE_PWD) as cursor:
                hora=datetime.datetime.now()
                consulta_v = f"SELECT TOP 1 FechaMarcado FROM RegistroMarcado WHERE PartNumber ='VERTICAL' AND TipoMarcado = 'INSPECCIÓN DE MARCADO' ORDER BY FechaMarcado DESC"
                cursor.execute(consulta_v)
                resultado_v = cursor.fetchone()[0]
                #dt = resultado_v.astype('datetime64[us]').astype(datetime)
                #ulltima_v = dt.timestamp()
                diff = (hora-resultado_v)
                #horas, restantes = divmod(diff, 3600)
                #minutos, segundos = divmod(restantes, 60)
                #minutos=int(minutos)
                #segundos_restantes = int(segundos_restantes)
                #timer = '{:02d}:{:02d}'.format(minutos, segundos_restantes) 
                #if horas>23:
                if diff.days>0:
                    vert.config(text='PENDIENTE',fg='red')
                else:
                    vert.config(text='OK',fg='black')
                consulta_h = f"SELECT TOP 1 FechaMarcado FROM RegistroMarcado WHERE PartNumber ='HORIZONTAL' AND TipoMarcado = 'INSPECCIÓN DE MARCADO' ORDER BY FechaMarcado DESC"
                cursor.execute(consulta_v)
                resultado_h = cursor.fetchone()[0]
                #dt = resultado_h.astype('datetime64[us]').astype(datetime)
                #ulltima_h = dt.timestamp()
                diff = (hora-resultado_h)
                #horas, restantes = divmod(diff, 3600)
                if diff.days>0:
                    hori.config(text='PENDIENTE',fg='red')
                else:
                    hori.config(text='OK',fg='black')
        except Exception as error:
            print (f'Error en la base de datos : ', error)
        time.sleep(1)

def registrar(operario, marcadora, estado, comentario,venta_inspec):
    if not operario or not marcadora or not estado:
        messagebox.showerror('ERROR', "FALTA ALGÚN CAMPO POR RELLENAR\nNº DE OPERARIO\nMARCADORA INSPECCIONADA\nESTADO DE MARCAJE")
        return
    else:
        print(operario,marcadora,comentario)
        try:
            with db_conn_po('ODBC Driver 17 for SQL Server', SISTE, SISTE_BD, SISTE_USR, SISTE_PWD) as cursor:
                hora=datetime.datetime.now()
                consulta2 = f'INSERT INTO RegistroMarcado (PartNumber, Orden, Operario, FechaMarcado, TipoMarcado, Notas) VALUES (?, ?, ?, ?, ?, ?)'
                cursor.execute(consulta2, marcadora, '000000', int(operario), hora, 'INSPECCIÓN DE MARCADO' , 'NOTAS: ' + str(comentario) )
        except Exception as error:
            print (f'Error en la base de datos : ', error)
        venta_inspec.destroy()

def inspeccion():
    char_limit=100

    def char_count(event):
        """This function allows typing up to the character limit and allows deletion"""
        count = len(text_comentarios.get('1.0', 'end-1c'))
        if count >= char_limit and event.keysym not in {'BackSpace', 'Delete'}:
            return 'break'  # dispose of the event, prevent typing

    venta_inspec = tk.Toplevel()
    venta_inspec.title("REGISTRO DE INSPECCIONES MARCADORAS NGVs")
    venta_inspec.geometry('600x150+300+300')
    label_operario = tk.Label(venta_inspec, text="Nº OPERARIO", font=('Arial',10,'bold'))
    label_operario.place(x=10,y=10)
    entry_operario = tk.Entry(venta_inspec, font='Arial',width=10)
    entry_operario.place(x=10,y=30)
    label_marcadora = tk.Label(venta_inspec, text="MARCADORA INSPECCIONADA", font=('Arial',10,'bold'))
    label_marcadora.place(x=130,y=10)
    combo_marcadora = ttk.Combobox(venta_inspec,values=['VERTICAL','HORIZONTAL'],width=30)
    combo_marcadora.place(x=130,y=30)
    label_marcaje = tk.Label(venta_inspec, text="ESTADO MARCAJE", font=('Arial',10,'bold'))
    label_marcaje.place(x=360,y=10)
    combo_marcaje = ttk.Combobox(venta_inspec,values=['OK','NO OK'],width=17)
    combo_marcaje.place(x=360,y=30)
    label_comentarios = tk.Label(venta_inspec, text="COMENTARIOS", font=('Arial',10,'bold'))
    label_comentarios.place(x=10,y=60)
    text_comentarios = tk.Text(venta_inspec, font='Arial', width=60, height=3)
    text_comentarios.place(x=10,y=80)
    text_comentarios.bind('<KeyPress>', char_count)
    text_comentarios.bind('<KeyRelease>', char_count)

    button_registrar = tk.Button(venta_inspec, text="REGISTRAR", font=('Arial',10,'bold'), command=lambda:registrar(entry_operario.get(),combo_marcadora.get(),combo_marcaje.get(),text_comentarios.get("1.0", tk.END),venta_inspec))
    button_registrar.bind("<FocusIn>", lambda event: venta_inspec.bind("<Return>", lambda event: button_registrar.invoke()))
    button_registrar.place(x=507,y=20)

    venta_inspec.focus_set()
    venta_inspec.mainloop()

def aprobado(KH,OF,REV):

    with db_conn_po('ODBC Driver 17 for SQL Server', SISTE, SISTE_BD, SISTE_USR, SISTE_PWD) as cursor:
    
        # COMPROBAR SI EL NUMERO DE SERIE DE FUNDIDO ESTÁ DUPLICADO DENTRO DE LAS OFS RESERVADAS, ABIERTAS O CERRADAS 
        consulta0 = f"SELECT Aprobado FROM ProgramaMarcado WITH (NOLOCK) WHERE Partnumber LIKE ? AND NumeroRevision LIKE ? AND EnUso LIKE '1'"
        #consulta0 = 'SELECT COUNT(*) FROM PARTMAT WHERE NSerieMaterial LIKE ?'
        cursor.execute(consulta0, KH, REV)
        resultado = cursor.fetchone()[0]
        if resultado == '0':
            messagebox.showerror('ERROR', "LA REVISIÓN MARCADA COMO EN USO NO ESTÁ APROBADA\nAVISAR A INGENIERIA ANTES DE MARCAR")
            open_login_window(KH,OF)
            return False
        else:
            return True

def doble_programa(KH):

    with db_conn_po('ODBC Driver 17 for SQL Server', SISTE, SISTE_BD, SISTE_USR, SISTE_PWD) as cursor:
    
        # COMPROBAR SI EL NUMERO DE SERIE DE FUNDIDO ESTÁ DUPLICADO DENTRO DE LAS OFS RESERVADAS, ABIERTAS O CERRADAS 
        consulta0 = f"SELECT COUNT(*) FROM ProgramaMarcado WITH (NOLOCK) WHERE Partnumber LIKE ? AND EnUso LIKE '1'"
        #consulta0 = 'SELECT COUNT(*) FROM PARTMAT WHERE NSerieMaterial LIKE ?'
        cursor.execute(consulta0, [KH])
        resultado = cursor.fetchone()[0]
        if resultado == 0:
            messagebox.showerror('ERROR', "NO HAY MARCAJE PARA ESTE PARTNUMBER\nAVISAR A INGENIERIA")
            return False
        elif resultado !=1:
            messagebox.showerror('ERROR', "DOS REVISIONES DISPONIBLES PARA ESTE PARTNUMBER\nAVISAR A INGENIERIA")
            return False
        else:
            return True

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
    button_login = tk.Button(login_window, text="INICIAR", command=lambda:(login(entry_username.get(),entry_password.get(),KH,OF),login_window.destroy()))
    button_login.bind("<FocusIn>", lambda event: login_window.bind("<Return>", lambda event: button_login.invoke()))
    button_login.place(x=15,y=110)
    login_window.wm_attributes("-topmost", True)
    login_window.focus_set()
    entry_username.focus_set()
    login_window.mainloop()

def login(user,pwd,KH,OF):
    username = user
    password = pwd

    if username == "REMARCAR" and password == "2022URRIA!":
        MAR, trama_regi, REV =marcar(KH,OF)
        time.sleep(40)
        if MAR==ACK:
            zero()
            registro_marcado(KH,OF,'NO OFICIAL',trama_regi,REV)
        else:
            messagebox.showerror('ERROR', "HA HABIDO UN PROBLEMA DURANTE EL MARCADO")
            zero()
        codigo_textbox.delete(0, tk.END)
        ventana.focus_set()
        codigo_textbox.focus_set()
    elif username == "ESPECIAL" and password == "2022URRIA!":
        MAR, trama_regi, REV =marcar(KH,OF)
        time.sleep(40)
        if MAR==ACK:
            zero()
            registro_marcado(KH,OF,'ESPECIAL',trama_regi,REV)
        else:
            messagebox.showerror('ERROR', "HA HABIDO UN PROBLEMA DURANTE EL MARCADO")
            zero()
        codigo_textbox.delete(0, tk.END)
        ventana.focus_set()
        codigo_textbox.focus_set()
    elif username == "HORIZONTAL" and password == "EGILE2023":
        try:
            with db_conn_po('ODBC Driver 17 for SQL Server', SISTE, SISTE_BD, SISTE_USR, SISTE_PWD) as cursor:
                consulta = 'SELECT PartNumber, Fichero FROM ProgramaMarcado WITH (NOLOCK) WHERE PartNumber LIKE ?'
                cursor.execute(consulta, 'CONTAHORI')
                CONTAHORI = cursor.fetchone()
                CONTAHORI_INT = int(CONTAHORI[1])
                hora=datetime.datetime.now()
                consulta = f'INSERT INTO RegistroMarcado (PartNumber, Orden, FechaMarcado, TipoMarcado) VALUES (?, ?, ?, ?)'
                cursor.execute(consulta, 'HORIZONTAL', CONTAHORI_INT, hora, 'CAMBIO DE PUNTA')
                CONTAHORI = str(0)
                consulta = 'UPDATE ProgramaMarcado SET Fichero = ? WHERE PartNumber=?'
                cursor.execute(consulta, CONTAHORI, 'CONTAHORI')

        except:
            messagebox.showerror('ERROR', "ERROR DE LA BASE DE DATOS")
            return
    elif username == "VERTICAL" and password == "EGILE2023":
        try:
            with db_conn_po('ODBC Driver 17 for SQL Server', SISTE, SISTE_BD, SISTE_USR, SISTE_PWD) as cursor:
                consulta = 'SELECT PartNumber, Fichero FROM ProgramaMarcado WITH (NOLOCK) WHERE PartNumber LIKE ?'
                cursor.execute(consulta, 'CONTAVERT')
                CONTAVERT = cursor.fetchone()
                hora=datetime.datetime.now()
                consulta2 = f'INSERT INTO RegistroMarcado (PartNumber, Orden, FechaMarcado, TipoMarcado) VALUES (?, ?, ?, ?)'
                CONTAVERT_INT = int(CONTAVERT[1])
                cursor.execute(consulta2, 'VERTICAL', CONTAVERT_INT, hora,'CAMBIO DE PUNTA')
                CONTAVERT = str(0)
                consulta = 'UPDATE ProgramaMarcado SET Fichero = ? WHERE PartNumber=?'
                cursor.execute(consulta, CONTAVERT, 'CONTAVERT')
        except:
            messagebox.showerror('ERROR', "ERROR DE LA BASE DE DATOS")
            return
    else:
        messagebox.showerror("Login", "Usuario o contraseña incorrectos")
        return

def registro_marcado(KH,OF,MARCA,trama_regi,REV):
    try:
        with db_conn_po('ODBC Driver 17 for SQL Server', SISTE, SISTE_BD, SISTE_USR, SISTE_PWD) as cursor:
                    # We truncate the table
                    #hora=time.time()
                    hora = datetime.datetime.now()
                    OFint=int(OF)
                    Registro = f'({KH}, {OFint}, {hora}, {MARCA})'
                    print(Registro)
                    consulta = f'INSERT INTO RegistroMarcado (PartNumber, Orden, FechaMarcado, TipoMarcado, ContenidoDM, Revision) VALUES (?, ?, ?, ?, ?, ?)'
                    cursor.execute(consulta, KH, OFint, hora, MARCA, trama_regi, REV)
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
        messagebox.showerror('ERROR', "ERROR AL IMPRIMIR LA HOJA DE REGISTRO\nCOMPRUEBA SI HAY ALGÚN EXCEL ABIERTO, CIÉRRALO Y VUELVE A INTENTARLO")

def ofids(coditex):
    SER_FUN = coditex.replace('SER', '')
    SER_FUN = SER_FUN.replace(' ', '')
    with db_conn_po('ODBC Driver 17 for SQL Server', BD, BD_BD, BD_USR, BD_PWD) as cursor:
        try:
            # COMPROBAR SI EL NUMERO DE SERIE DE FUNDIDO ESTÁ DUPLICADO DENTRO DE LAS OFS RESERVADAS, ABIERTAS O CERRADAS 
            
            consulta0 = f"SELECT PARTMAT.NumeroOrden, PARTMAT.CodigoComponente, ORDEN.Situacion FROM PARTMAT WITH (NOLOCK) INNER JOIN ORDEN WITH (NOLOCK) ON PARTMAT.NumeroOrden = ORDEN.Numero WHERE (PARTMAT.NSerieMaterial LIKE ?)"
            #consulta0 = 'SELECT COUNT(*) FROM PARTMAT WHERE NSerieMaterial LIKE ?'
            cursor.execute(consulta0, [SER_FUN])
            resultado = cursor.fetchone()
            OF = str(resultado[0])
            Codi_Art = str(resultado[1])
            situacion = str(resultado[2]).upper()
            if situacion == 'RESERVADA':
                #messagebox.showerror('ERROR', "ESTA OF ESTÁ EN SITUACIÓN RESERVADA\nINTENTA SACAR LA OF DE FORMA NORMAL")
                soloOF(coditex)
            else:
                logo(Codi_Art)
                hwnd = win32gui.FindWindow(None, 'Gestor de Aplicaciones Comerciales e Industriales')
                #ventanaIDS = pygetwindow.getWindowsWithTitle('Inicio de trabajo OF')
                if hwnd == 0:
                    shortcut_path = r"C:\Users\taller20\Desktop\IDSWIN.lnk"
                    os.startfile(shortcut_path)
                    time.sleep(7)
                elif hwnd != 0:
                    win32gui.ShowWindow(hwnd, 5)
                    win32gui.SetForegroundWindow(hwnd)

                pyautogui.click(x=800, y=450)
                time.sleep(12)
                pyautogui.click(x=150, y=40)
                time.sleep(0.7)
                pyautogui.click(x=150, y=270)                      
                time.sleep(0.7)
                pyautogui.click(x=900, y=340)
                time.sleep(0.7)   
                pyautogui.click(x=900, y=380)          
                pyperclip.copy(OF)
                pyautogui.hotkey('tab')
                pyautogui.hotkey('tab')
                pyautogui.hotkey('tab')
                pyautogui.hotkey('ctrl', 'v')
                pyautogui.hotkey('tab')
                pyperclip.copy(OF)
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
                #time.sleep(6)
                #pyautogui.click(x=500, y=50)
                #pyautogui.hotkey('tab')
                #pyautogui.hotkey('tab')
                #pyautogui.hotkey('tab')
                #pyautogui.hotkey('tab')
                #pyautogui.hotkey('tab')
                #pyautogui.hotkey('enter')
        except:
            messagebox.showerror('ERROR', "ERROR AL INTENTAR IMPRIMIR LA OF\nAVISA A INGENIERÍA")
            
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
            print(archivo)
            logo(archivo)
            shortcut_path = r"C:\Users\taller20\Desktop\PISTOLA.lnk"
            os.startfile(shortcut_path)
            time.sleep(5)
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

        #if hex_string == ACK:
        #    resultado_label2.config(text='Reiniciando Marcadora')
        #else:
        #    resultado_label2.config(text='Error Reset')
    #else:
    #        print('No hay respuesta')
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
        with db_conn_po('ODBC Driver 17 for SQL Server', SISTE, SISTE_BD, SISTE_USR, SISTE_PWD) as cursor:
            # We truncate the table
            consulta = "SELECT PartNumber, Fichero, Marcadora, TipoMarcado, MFR, SER, PNR, ASSY, NumeroRevision FROM ProgramaMarcado WITH (NOLOCK) WHERE PartNumber LIKE ? AND EnUso ='1'"
            cursor.execute(consulta, [KH])
            resultado = cursor.fetchone() #ESTA FUNCIÓN SACA LA OF, EL CÓDIGO ARTÍCULO Y EL SN DE FUNDIDO
            Fichero=resultado[1]
            MARCADORA=resultado[2]
            Tipo_Marcado=str(resultado[3])
            MFR = resultado[4]
            SER = resultado[5]
            PNR = resultado[6]
            ASSY=resultado[7]    
            REV=resultado[8]


            # CONTADOR DE MARCAJES 
            if KH in ['KH18990', 'KH18995']:
                consulta = 'SELECT PartNumber, Fichero FROM ProgramaMarcado WITH (NOLOCK) WHERE PartNumber LIKE ?'
                cursor.execute(consulta, 'CONTAHORI')
                resultado = cursor.fetchone() #ESTA FUNCIÓN SACA LA OF, EL CÓDIGO ARTÍCULO Y EL SN DE FUNDIDO
                CONTAHORI = int(resultado[1])
                CONTAHORI += 1
                consulta = 'SELECT PartNumber, Fichero FROM ProgramaMarcado WITH (NOLOCK) WHERE PartNumber LIKE ?'
                cursor.execute(consulta, 'MAXHORI')
                resultado = cursor.fetchone() #ESTA FUNCIÓN SACA LA OF, EL CÓDIGO ARTÍCULO Y EL SN DE FUNDIDO
                MAXHORI = int(resultado[1])
                if CONTAHORI >= MAXHORI:
                    thread_control_cambio = threading.Thread(target=messagebox.showerror, args=('CAMBIO DE PUNTA', "CAMBIAR LA PUNTA EN MARCADORA HORIZONTAL\n¡ATENCIÓN!\nCOMPRUEBA QUE LA PIEZA SE HA MARCADO CORRECTAMENTE"))
                    thread_control_cambio.start()
                    #messagebox.showerror('CAMBIO DE PUNTA', "CAMBIAR LA PUNTA EN MARCADORA HORIZONTAL\n¡ATENCIÓN! PIEZA NO MARCADA, ACEPTAR PARA MARCAR")
                CONTAHORI = str(CONTAHORI)
                consulta = 'UPDATE ProgramaMarcado SET Fichero = ? WHERE PartNumber=?'
                cursor.execute(consulta, CONTAHORI, 'CONTAHORI')
            
            else:
                consulta = 'SELECT PartNumber, Fichero FROM ProgramaMarcado WITH (NOLOCK) WHERE PartNumber LIKE ?'
                cursor.execute(consulta, 'CONTAVERT')
                resultado = cursor.fetchone() #ESTA FUNCIÓN SACA LA OF, EL CÓDIGO ARTÍCULO Y EL SN DE FUNDIDO
                CONTAVERT = int(resultado[1])
                CONTAVERT += 1
                consulta = 'SELECT PartNumber, Fichero FROM ProgramaMarcado WITH (NOLOCK) WHERE PartNumber LIKE ?'
                cursor.execute(consulta, 'MAXVERT')
                resultado = cursor.fetchone() #ESTA FUNCIÓN SACA LA OF, EL CÓDIGO ARTÍCULO Y EL SN DE FUNDIDO
                MAXVERT = int(resultado[1])
                if CONTAVERT >= MAXVERT:
                    thread_control_cambio = threading.Thread(target=messagebox.showerror, args=('CAMBIO DE PUNTA', "CAMBIAR LA PUNTA EN MARCADORA VERTICAL\n¡ATENCIÓN!\nCOMPRUEBA QUE LA PIEZA SE HA MARCADO CORRECTAMENTE"))
                    thread_control_cambio.start()
                    #messagebox.showerror('CAMBIO DE PUNTA', "CAMBIAR LA PUNTA EN MARCADORA VERTICAL\n¡ATENCIÓN! PIEZA NO MARCADA, ACEPTAR PARA MARCAR")
                CONTAVERT = str(CONTAVERT)
                consulta = 'UPDATE ProgramaMarcado SET Fichero = ? WHERE PartNumber=?'
                cursor.execute(consulta, CONTAVERT, 'CONTAVERT')

    except:
        messagebox.showerror('ERROR', "NO HAY UN PROGRAMA DEFINIDO PARA ESTA PIEZA")
        return
    
    if aprobado(KH,OF,REV) != True:
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

    '''
    if hex_string == ACK:
        resultado_label2.config(text='Programa cargado y Marcadora lista')
    elif hex_string == NACK:
        resultado_label2.config(text='Marcadora no disponible')
    else:
        resultado_label2.config(text='ERROR Marcado')
        '''

    client_socket.close()

    # cierra la conexión con el dispositivo

    # MARCAR TEXTO

    LT = '0c'

    SER_txt = SER + '201293DMP' + OF
    SER_txt_hex = ' '.join([hex(ord(c))[2:].zfill(2) for c in SER_txt])
    SER_dm = 'SER 201293DMP' + OF
    SER_dm_hex = ' '.join([hex(ord(c))[2:].zfill(2) for c in SER_dm])

    PNR_txt = PNR + KH
    PNR_txt_hex = ' '.join([hex(ord(c))[2:].zfill(2) for c in PNR_txt])
    PNR_dm = 'PNR ' + KH
    PNR_dm_hex = ' '.join([hex(ord(c))[2:].zfill(2) for c in PNR_dm])

    MFR_txt = MFR + 'K0680'
    MFR_txt_hex = ' '.join([hex(ord(c))[2:].zfill(2) for c in MFR_txt])
    MFR_dm = 'MFR ' + 'K0680'
    MFR_dm_hex = ' '.join([hex(ord(c))[2:].zfill(2) for c in MFR_dm])

    if ASSY=='ASSY':
        ASSY_hex = ' '.join([hex(ord(c))[2:].zfill(2) for c in ASSY])
    else:
        ASSY_hex='20'
    
    DM = Corchete + ' ' + RS + ' ' + Doce + ' ' + GS + ' ' + MFR_dm_hex + ' ' + GS + ' ' + SER_dm_hex + ' ' + GS + ' ' + PNR_dm_hex + ' ' + RS + ' ' + EOT


    for i in range(0, 2):

        if MARCADORA=='VERTICAL':
            Marcar_Texto = '10 02 ' + LT + ' 00 36'

        elif MARCADORA=='HORIZONTAL':
            Marcar_Texto = '10 02 ' + LT + ' 00 36' + ' ' + Activar_Salida_1
        
        if Tipo_Marcado == '1':#tipo de marcado con DM único y rectangular

            trama = Marcar_Texto + ' ' + Punto + ' 0a ' \
                    + Data_Matrix_Rect_abre + ' ' + DM + ' ' + Data_Matrix_Rect_cierra + ' 0a ' \
                    + MFR_txt_hex + ' 0a ' \
                    + SER_txt_hex + ' 0a ' \
                    + '0a ' \
                    + PNR_txt_hex + ' 0a ' \
                    + ASSY_hex
            trama_regi = Corchete_regi + RS_regi + '12' + GS_regi + MFR_dm + GS_regi + SER_dm + GS_regi + PNR_dm + RS_regi + EOT_regi

        elif Tipo_Marcado == '2':#tipo de marcado con DM único y cradrado

            trama = Marcar_Texto + ' ' + Punto + ' 0a ' \
                    + DM_SQ_abre + ' ' + DM + ' ' + DM_SQ_cierra + ' 0a ' \
                    + MFR_txt_hex + ' 0a ' \
                    + SER_txt_hex + ' 0a ' \
                    + '0a ' \
                    + PNR_txt_hex + ' 0a ' \
                    + ASSY_hex
            trama_regi = Corchete_regi + RS_regi + '12' + GS_regi + MFR_dm + GS_regi + SER_dm + GS_regi + PNR_dm + RS_regi + EOT_regi
            
        elif Tipo_Marcado=='3':#tipo de marcado antiguo, con 2 DM cradrados
            DM1 = MFR_dm_hex + ' 20 ' + SER_dm_hex
            DM2 = PNR_dm_hex
            trama = Marcar_Texto + ' ' + Punto + ' 0a ' \
                    + DM_SQ_abre + ' ' + DM1 + ' ' + DM_SQ_cierra + ' 0a ' \
                    + MFR_txt_hex + ' 0a ' \
                    + SER_txt_hex + ' 0a ' \
                    + DM_SQ_abre + ' ' + DM2 + ' ' + DM_SQ_cierra + ' 0a ' \
                    + PNR_txt_hex + ' 0a ' \
                    + ASSY_hex
            trama_regi = MFR_txt + SER_txt + PNR_txt
            
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
        #resultado_label2.config(text='Texto transmitido, inicio de marcado...')
        print('Marcando...')
        client_socket.close()
        codigo_textbox.insert(0, '')  
        return ACK, trama_regi, REV
    else:
        #resultado_label2.config(text='Error en la trama')
        client_socket.close()
        return 0
    
def no_oficial(coditex):
    
    SER_FUN = coditex.replace('SER', '')
    SER_FUN = SER_FUN.replace(' ', '')

    with db_conn_po('ODBC Driver 17 for SQL Server', BD, BD_BD, BD_USR, BD_PWD) as cursor:
        try:
            # COMPROBAR SI EL NUMERO DE SERIE DE FUNDIDO ESTÁ DUPLICADO DENTRO DE LAS OFS RESERVADAS, ABIERTAS O CERRADAS 
            consulta0 = f"SELECT COUNT(*) FROM PARTMAT WITH (NOLOCK) INNER JOIN ORDEN WITH (NOLOCK) ON PARTMAT.NumeroOrden = ORDEN.Numero WHERE (PARTMAT.NSerieMaterial LIKE ?)"
            #consulta0 = 'SELECT COUNT(*) FROM PARTMAT WHERE NSerieMaterial LIKE ?'
            cursor.execute(consulta0, [SER_FUN])
            resultado = cursor.fetchone()[0]

            if resultado==1:
                #SI SOLO HAY UNA, EXTRAER OF, CÓDIGO ARTÍCULO Y PARTNUMBER 
                # EN ESTE CASO SE PUEDEN MARCAR PIEZAS ENTRE LAS RESERVADAS Y LAS ABIERTAS
                consulta1 = f"SELECT PARTMAT.NumeroOrden, PARTMAT.CodigoComponente FROM PARTMAT WITH (NOLOCK) INNER JOIN ORDEN WITH (NOLOCK) ON PARTMAT.NumeroOrden = ORDEN.Numero WHERE (PARTMAT.NSerieMaterial LIKE ?)"
                cursor.execute(consulta1, [SER_FUN])
                resultado = cursor.fetchone()
                OF=str(resultado[0])
                Codi_Art=str(resultado[1])
                consulta2='SELECT CodigoAlternativo FROM ARTICULO WITH (NOLOCK) WHERE CODIGO LIKE ?'
                cursor.execute(consulta2, [Codi_Art])
                resultado2 = cursor.fetchone()
                KH = resultado2[0]
                        
            elif resultado >= 2:
                PN_FUN=part_rep()

                consulta3=f"SELECT Codigo FROM ARTICULO WITH (NOLOCK) WHERE CodigoAlternativo LIKE ?"
                cursor.execute(consulta3, [PN_FUN])
                resultado3 = cursor.fetchone()
                CO_MA = resultado3[0]

                consulta1 = f"SELECT PARTMAT.NumeroOrden, PARTMAT.CodigoComponente FROM PARTMAT WITH (NOLOCK) INNER JOIN ORDEN WITH (NOLOCK) ON PARTMAT.NumeroOrden = ORDEN.Numero WHERE (PARTMAT.CodigoMaterial={CO_MA}) AND (PARTMAT.NSerieMaterial LIKE ?)"
                cursor.execute(consulta1, [SER_FUN])
                resultado = cursor.fetchone()
                OF=str(resultado[0])
                Codi_Art=str(resultado[1])
                consulta2='SELECT CodigoAlternativo FROM ARTICULO WITH (NOLOCK) WHERE CODIGO LIKE ?'
                cursor.execute(consulta2, [Codi_Art])
                resultado2 = cursor.fetchone()
                KH = resultado2[0]

        except:
            messagebox.showerror('ERROR', "NO SE ENCUENTRA EL NÚMERO DE SERIE\nAVISA A INGENIERÍA")
            return
    if doble_programa(KH) != True:
        return
    open_login_window(KH,OF)

    estado()
    cursor.close()  
    ventana.focus()
    codigo_textbox.focus_set()

def soloHOJA(coditex):

    SER_FUN = coditex.replace('SER', '')
    SER_FUN = SER_FUN.replace(' ', '')

    with db_conn_po('ODBC Driver 17 for SQL Server', BD, BD_BD, BD_USR, BD_PWD) as cursor:
        try:
            # COMPROBAR SI EL NUMERO DE SERIE DE FUNDIDO ESTÁ DUPLICADO DENTRO DE LAS OFS RESERVADAS, ABIERTAS O CERRADAS 
            consulta0 = f"SELECT COUNT(*) FROM PARTMAT WITH (NOLOCK) INNER JOIN ORDEN WITH (NOLOCK) ON PARTMAT.NumeroOrden = ORDEN.Numero WHERE (PARTMAT.NSerieMaterial LIKE ?)"
            #consulta0 = 'SELECT COUNT(*) FROM PARTMAT WHERE NSerieMaterial LIKE ?'
            cursor.execute(consulta0, [SER_FUN])
            resultado = cursor.fetchone()[0]

            if resultado==1:
                #SI SOLO HAY UNA, EXTRAER OF, CÓDIGO ARTÍCULO Y PARTNUMBER 
                # EN ESTE CASO SE PUEDEN MARCAR PIEZAS ENTRE LAS RESERVADAS Y LAS ABIERTAS
                consulta1 = f"SELECT PARTMAT.NumeroOrden, PARTMAT.CodigoComponente FROM PARTMAT WITH (NOLOCK) INNER JOIN ORDEN WITH (NOLOCK) ON PARTMAT.NumeroOrden = ORDEN.Numero WHERE (PARTMAT.NSerieMaterial LIKE ?)"
                cursor.execute(consulta1, [SER_FUN])
                resultado = cursor.fetchone()
                OF=str(resultado[0])
                Codi_Art=str(resultado[1])
                consulta2='SELECT CodigoAlternativo FROM ARTICULO WITH (NOLOCK) WHERE CODIGO LIKE ?'
                cursor.execute(consulta2, [Codi_Art])
                resultado2 = cursor.fetchone()
                KH = resultado2[0]
                        
            elif resultado >= 2:
                PN_FUN=part_rep()

                consulta3=f"SELECT Codigo FROM ARTICULO WITH (NOLOCK) WHERE CodigoAlternativo LIKE ?"
                cursor.execute(consulta3, [PN_FUN])
                resultado3 = cursor.fetchone()
                CO_MA = resultado3[0]

                consulta1 = f"SELECT PARTMAT.NumeroOrden, PARTMAT.CodigoComponente FROM PARTMAT WITH (NOLOCK) INNER JOIN ORDEN WITH (NOLOCK) ON PARTMAT.NumeroOrden = ORDEN.Numero WHERE (PARTMAT.CodigoMaterial={CO_MA}) AND (PARTMAT.NSerieMaterial LIKE ?)"
                cursor.execute(consulta1, [SER_FUN])
                resultado = cursor.fetchone()
                OF=str(resultado[0])
                Codi_Art=str(resultado[1])
                consulta2='SELECT CodigoAlternativo FROM ARTICULO WITH (NOLOCK) WHERE CODIGO LIKE ?'
                cursor.execute(consulta2, [Codi_Art])
                resultado2 = cursor.fetchone()
                KH = resultado2[0]

        except:
            messagebox.showerror('ERROR', "NO SE ENCUENTRA EL NÚMERO DE SERIE\nAVISA A INGENIERÍA")
            return

    imprimir_excel(OF,KH)

def soloOF(coditex):

    SER_FUN = coditex.replace('SER', '')
    SER_FUN = SER_FUN.replace(' ', '')

    with db_conn_po('ODBC Driver 17 for SQL Server', BD, BD_BD, BD_USR, BD_PWD) as cursor:
        try:
            # COMPROBAR SI EL NUMERO DE SERIE DE FUNDIDO ESTÁ DUPLICADO DENTRO DE LAS OFS RESERVADAS, ABIERTAS O CERRADAS 
            consulta0 = f"SELECT COUNT(*) FROM PARTMAT WITH (NOLOCK) INNER JOIN ORDEN WITH (NOLOCK) ON PARTMAT.NumeroOrden = ORDEN.Numero WHERE (PARTMAT.NSerieMaterial LIKE ?)"
            #consulta0 = 'SELECT COUNT(*) FROM PARTMAT WHERE NSerieMaterial LIKE ?'
            cursor.execute(consulta0, [SER_FUN])
            resultado = cursor.fetchone()[0]

            if resultado==1:
                #SI SOLO HAY UNA, EXTRAER OF, CÓDIGO ARTÍCULO Y PARTNUMBER 
                # EN ESTE CASO SE PUEDEN MARCAR PIEZAS ENTRE LAS RESERVADAS Y LAS ABIERTAS
                consulta1 = f"SELECT PARTMAT.NumeroOrden, PARTMAT.CodigoComponente FROM PARTMAT WITH (NOLOCK) INNER JOIN ORDEN WITH (NOLOCK) ON PARTMAT.NumeroOrden = ORDEN.Numero WHERE (PARTMAT.NSerieMaterial LIKE ?)"
                cursor.execute(consulta1, [SER_FUN])
                resultado = cursor.fetchone()
                OF=str(resultado[0])
                Codi_Art=str(resultado[1])
                consulta2='SELECT CodigoAlternativo FROM ARTICULO WITH (NOLOCK) WHERE CODIGO LIKE ?'
                cursor.execute(consulta2, [Codi_Art])
                resultado2 = cursor.fetchone()
                KH = resultado2[0]
                        
            elif resultado >= 2:
                PN_FUN=part_rep()

                consulta3=f"SELECT Codigo FROM ARTICULO WITH (NOLOCK) WHERE CodigoAlternativo LIKE ?"
                cursor.execute(consulta3, [PN_FUN])
                resultado3 = cursor.fetchone()
                CO_MA = resultado3[0]

                consulta1 = f"SELECT PARTMAT.NumeroOrden, PARTMAT.CodigoComponente FROM PARTMAT WITH (NOLOCK) INNER JOIN ORDEN WITH (NOLOCK) ON PARTMAT.NumeroOrden = ORDEN.Numero WHERE (PARTMAT.CodigoMaterial={CO_MA}) AND (PARTMAT.NSerieMaterial LIKE ?)"
                cursor.execute(consulta1, [SER_FUN])
                resultado = cursor.fetchone()
                OF=str(resultado[0])
                Codi_Art=str(resultado[1])
                consulta2='SELECT CodigoAlternativo FROM ARTICULO WITH (NOLOCK) WHERE CODIGO LIKE ?'
                cursor.execute(consulta2, [Codi_Art])
                resultado2 = cursor.fetchone()
                KH = resultado2[0]

        except:
            messagebox.showerror('ERROR', "NO SE ENCUENTRA EL NÚMERO DE SERIE\nAVISA A INGENIERÍA")
            return

    IDS=buscar_y_pegar('SER ' + SER_FUN,Codi_Art)#Abre la OF y la Imprime, necesita el SER y el codigo de articulo para el logo de color
    if IDS==0:
        return


def oficial(coditex):

    SER_FUN = coditex.replace('SER', '')
    SER_FUN = SER_FUN.replace(' ', '')

    #if '16720ZC2' in SER_FUN:
    #    messagebox.showerror('ERROR', "Nº DE SERIE EN CUARENTENA\nMARCA OTRA PIEZA Y AVISA A INGENIERÍA")
    #    return

    with db_conn_po('ODBC Driver 17 for SQL Server', BD, BD_BD, BD_USR, BD_PWD) as cursor:
        try:
            # COMPROBAR SI EL NUMERO DE SERIE DE FUNDIDO ESTÁ DUPLICADO DENTRO DE LAS OFS RESERVADAS
            consulta0 = f"SELECT COUNT(*) FROM PARTMAT WITH (NOLOCK) INNER JOIN ORDEN WITH (NOLOCK) ON PARTMAT.NumeroOrden = ORDEN.Numero WHERE (PARTMAT.NSerieMaterial LIKE ?)"
            #consulta0 = 'SELECT COUNT(*) FROM PARTMAT WHERE NSerieMaterial LIKE ?'
            cursor.execute(consulta0, [SER_FUN])
            resultado = cursor.fetchone()[0]

            if resultado==1:
                #SI SOLO HAY UNA, EXTRAER OF, CÓDIGO ARTÍCULO Y PARTNUMBER DENTRO DE LAS RESERVADAS
                consulta1 = f"SELECT PARTMAT.NumeroOrden, PARTMAT.CodigoComponente, ORDEN.Situacion FROM PARTMAT WITH (NOLOCK) INNER JOIN ORDEN WITH (NOLOCK) ON PARTMAT.NumeroOrden = ORDEN.Numero WHERE (PARTMAT.NSerieMaterial LIKE ?)"
                cursor.execute(consulta1, [SER_FUN])
                resultado = cursor.fetchone()
                OF=str(resultado[0])
                Codi_Art=str(resultado[1])
                situacion = str(resultado[2]).upper()
                if situacion != 'RESERVADA':
                    messagebox.showerror('ERROR', f"LA OF ESTÁ EN SITUACIÓN {situacion}\nAVISAR A INGENIERÍA")
                    return
                consulta2=f"SELECT CodigoAlternativo FROM ARTICULO WITH (NOLOCK) WHERE CODIGO LIKE ?"
                cursor.execute(consulta2, [Codi_Art])
                resultado2 = cursor.fetchone()
                KH = resultado2[0]
                        
            elif resultado >= 2:
                PN_FUN=part_rep()

                consulta3=f"SELECT Codigo FROM ARTICULO WITH (NOLOCK) WHERE CodigoAlternativo LIKE ?"
                cursor.execute(consulta3, [PN_FUN])
                resultado3 = cursor.fetchone()
                CO_MA = resultado3[0]

                consulta1 = f"SELECT PARTMAT.NumeroOrden, PARTMAT.CodigoComponente FROM PARTMAT WITH (NOLOCK) INNER JOIN ORDEN WITH (NOLOCK) ON PARTMAT.NumeroOrden = ORDEN.Numero WHERE (PARTMAT.CodigoMaterial={CO_MA}) AND (PARTMAT.NSerieMaterial LIKE ?)"
                cursor.execute(consulta1, [SER_FUN])
                resultado = cursor.fetchone()
                OF=str(resultado[0])
                Codi_Art=str(resultado[1])
                consulta2='SELECT CodigoAlternativo FROM ARTICULO WITH (NOLOCK) WHERE CODIGO LIKE ?'
                cursor.execute(consulta2, [Codi_Art])
                resultado2 = cursor.fetchone()
                KH = resultado2[0]
            elif resultado==0:
                messagebox.showerror(f'ERROR', "OF {OF}")
                return

        except:
            messagebox.showerror('ERROR', "NO SE ENCUENTRA EL NÚMERO DE SERIE")
            return
        
        consulta = 'SELECT XLIPA.Descripcion FROM XLIPA WITH (NOLOCK) INNER JOIN ORDEN WITH (NOLOCK) ON XLIPA.NumeroIPA = ORDEN.NumeroIPA AND XLIPA.Nivel = ORDEN.NivelIPA AND XLIPA.Revision = ORDEN.RevisionIPA WHERE ORDEN.Numero LIKE ?'
        
        cursor.execute(consulta, [OF])
        resultado = cursor.fetchone()
        desc = resultado[0]
        if SER_FUN[0].isalpha() and 'LPO' not in desc:
            messagebox.showerror('ERROR', "EL IPA DE ESTA PIEZA NO ES CORRECTO")
            return       

    #COMPROBAR SI HAY REGISTRO SELECT EXISTS(SELECT * FROM nombre_de_la_tabla WHERE nombre_de_la_columna = valor_buscado);
    #try:
    with db_conn_po('ODBC Driver 17 for SQL Server', SISTE, SISTE_BD, SISTE_USR, SISTE_PWD) as cursor:
        consulta = 'SELECT COUNT(*) FROM RegistroMarcado WITH (NOLOCK) WHERE Orden LIKE ?'
        cursor.execute(consulta, [OF])
        resultado = cursor.fetchone()[0]
    if resultado!=0:
        messagebox.showerror('ERROR', "ESTA PIEZA YA SE HA MARCADO")
        return

    #COMPROBAR SI HAY MÁS DE UN PROGRAMA ACTIVADO 
    if doble_programa(KH) != True:
        return
    
    IDS=buscar_y_pegar('SER ' + SER_FUN,Codi_Art)#Abre la OF y la Imprime, necesita el SER y el codigo de articulo para el logo de color
    if IDS==0:
        return
    
    #IMPRIMIR HOJA DE REGISTRO

    imprimir_excel(OF,KH)

    MAR, trama_regi, REV =marcar(KH,OF)
    time.sleep(40)
    if MAR==ACK:
        hwnd = ventana.winfo_id()
        win32gui.SetForegroundWindow(hwnd)
        codigo_textbox.delete(0, tk.END)
        #resultado_label2.config(text='Marcado terminado')
        zero()
        registro_marcado(KH,OF,'OFICIAL',trama_regi,REV)
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
boton_reset.bind("<FocusIn>", lambda event: boton_reset.bind("<Return>", lambda event=None: boton_reset.invoke()))

# Botón de cerrar
boton_cerrar = tk.Button(ventana, text='INSPECCIÓN', width=10, height=2, command=lambda: inspeccion())
boton_cerrar.place(x=300, y=155)

# Botón SOLO OF
boton_OF = tk.Button(ventana, text='OF', width=12, height=1, command=lambda: ofids(codigo_textbox.get()))
boton_OF.place(x=195, y=150)
boton_OF.bind("<FocusIn>", lambda event: boton_OF.bind("<Return>", lambda event=None: boton_OF.invoke()))


# Botón SOLO HOJA DE REGISTRO
boton_HR = tk.Button(ventana, text='HOJA REGISTRO', width=12, height=1, command=lambda: soloHOJA(codigo_textbox.get()))
boton_HR.place(x=195, y=180)

# label 1
resultado_label1 = tk.Label(ventana, text='SER ->', font=('Arial', 20, 'bold'), justify='left')
resultado_label1.place(x=80, y=110, anchor=tk.CENTER)

# Crear la caja de texto
codigo_textbox = tk.Entry(ventana, font=('Arial', 20, 'bold'), justify='center')
codigo_textbox.place(x=250, y=110, width=250, height=50, anchor=tk.CENTER)
codigo_textbox.focus_set()
codigo_textbox.bind("<FocusIn>", lambda event: codigo_textbox.bind('<Return>', lambda event: oficial(codigo_textbox.get())))

# label 2
#resultado_label2 = tk.Label(ventana, text='', font=('Arial', 10), justify='left')
#resultado_label2.place(x=250, y=220, anchor=tk.CENTER)

# Cargar la imagen utilizando opencv
imagen1 = Image.open("M:\\OLANET\\G02-13 MITUTOYO CRYSTA APEX\\NGVLEAN\\scripts\\Logo_Egile_Mechanics.jpg")
imagen1 = imagen1.resize((128, 73))
photo = ImageTk.PhotoImage(imagen1)

# label 4
resultado_label_egile = tk.Label(ventana, image=photo, relief=tk.RAISED)
resultado_label_egile.place(x=70, y=40, anchor=tk.CENTER)

#VERIFICACIÓN ESTADO MARCAJE

INSPE_LABEL = tk.Label(ventana,text='INSPECCIÓN ESTADO DE MARCAJE', font=('Arial', 10, 'bold'),justify='left')
INSPE_LABEL.place(x=20, y=260, anchor=tk.SW)

INSPE_LABEL_VERT = tk.Label(ventana,text='VERTICAL->', font=('Arial', 10, 'bold'),justify='left')
INSPE_LABEL_VERT.place(x=260, y=245, anchor=tk.SW)

INSPE_TIMER_VERT = tk.Label(ventana,text='PENDIENTE', font=('Arial', 10, 'bold'), fg='red',justify='left')
INSPE_TIMER_VERT.place(x=355, y=245, anchor=tk.SW)

INSPE_LABEL_HORI = tk.Label(ventana,text='HORIZONTAL->', font=('Arial', 10, 'bold'),justify='left')
INSPE_LABEL_HORI.place(x=250, y=275, anchor=tk.SW)

INSPE_TIMER_HORI = tk.Label(ventana,text='OK', font=('Arial', 10, 'bold'), fg='black',justify='left')
INSPE_TIMER_HORI.place(x=355, y=275, anchor=tk.SW)

thread_alineamiento = threading.Thread(target=control, args=(INSPE_TIMER_VERT, INSPE_TIMER_HORI))
thread_alineamiento.start()

ventana.mainloop()
