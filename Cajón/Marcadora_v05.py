import socket
import tkinter as tk
import tkinter.messagebox as messagebox
from PIL import ImageTk, Image
import time
import re
import select
import pyodbc as po

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
Activar_Marcadora_Horizontal = '1e 4d 53 31 41 1f'

NACK = '10 00 07 00 ff 9f 3a'
ACK = '10 00 07 00 00 81 ca'
RESPUESTA_MARCHA = '10 00 08 00 32 00 FC B0'

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

    # espera a recibir la respuesta del dispositivo
    response_bytes = client_socket.recv(1024)

    hex_string = ' '.join(['{:04x}'.format(byte) for byte in
                          response_bytes])  # convertir cada byte en su representación hexadecimal y unirlos en una cadena

    if hex_string == '10 00 08 00 03 00 ca 14':
        resultado_label2.config(text='Marcadora lista y a la espera')
    else:
        resultado_label2.config(text='Marcadora Ocupada')

    # cierra la conexión con el dispositivo
    client_socket.close()


def marcar(coditex):
    # CARGAR EL FICHERO QUE CORRESPONDA
    SER_FUN = coditex.replace('SER ', '')


    #COMPROBAR SI EXISTE EL SER
    #COMPROBAR SI ESTÁ DUPLICADO
        #EN CASO DE QUE ESTÉ DUPLICADO CONSULTAR QUÉ PARTNUMBER ES EL QUE SE QUIERE MARCAR

    with db_conn_po('ODBC Driver 17 for SQL Server', 'MENITBD', 'IDSDMP1', 'lecturaIDS', '0') as cursor:
        try:
            # COMPROBAR SI EL NUMERO DE SERIE DE FUNDIDO ESTÁ DUPLICADO
            consulta0 = 'SELECT COUNT(*) FROM PARTMAT WHERE NSERIEMATERIAL LIKE ?'
            cursor.execute(consulta0, [SER_FUN])
            resultado = cursor.fetchone()[0]

            if resultado==1:
                consulta1 = 'SELECT NumeroOrden,CodigoComponente,NSerieMaterial FROM PARTMAT WHERE NSERIEMATERIAL LIKE ?'
                cursor.execute(consulta1, [SER_FUN])
                resultado = cursor.fetchone() #ESTA FUNCIÓN SACA LA OF, EL CÓDIGO ARTÍCULO Y EL SN DE FUNDIDO
                OF=str(resultado[0])
                Codi_Art=str(resultado[1])
                consulta2='SELECT CodigoAlternativo FROM ARTICULO WHERE CODIGO LIKE ?'
                cursor.execute(consulta2, [Codi_Art])
                resultado2 = cursor.fetchone()
                KH = resultado2[0]
                print(KH)
                consulta3 ='SELECT Situacion FROM ORDEN WHERE Numero LIKE ?'
                cursor.execute(consulta2, [OF])
                resultado3 = cursor.fetchone()
                if resultado3=='Abierta':
                    messagebox.showerror('ERROR', "ESTA OF NO ESTÁ RESERVADA")
                    return
                elif resultado3=='Cerrada':
                    messagebox.showerror('ERROR', "ESTA OF NO ESTÁ RESERVADA")
                    return

            elif resultado >= 2:
                messagebox.showerror('ERROR', "ESTE NÚMERO DE SERIE YA SE HA USADO \n APARTA LA PIEZA Y AVISA A INGENIERÍA")
                return
        except:
            messagebox.showerror('ERROR', "NO SE ENCUENTRA EL NÚMERO DE SERIE")
            return
    
    #COMPROBAR SI HAY REGISTRO SELECT EXISTS(SELECT * FROM nombre_de_la_tabla WHERE nombre_de_la_columna = valor_buscado);
    try:
        with db_conn_po('ODBC Driver 17 for SQL Server', 'MENITSISTEPLANT', 'marcadongvs', 'landerurgoiti', '2023Urria!') as cursor:
            consulta = 'SELECT COUNT(*) FROM RegistroMarcado WHERE Orden LIKE ?'
            cursor.execute(consulta, [OF])
            resultado = cursor.fetchone()[0]
        if resultado!=0:
            

    except:
        messagebox.showerror('ERROR', "ERROR EN LA COMPROBACIÓN DEL REGISTRO \n AVISAR A INGENIERÍA")








    #COMPROBAR OF RESERVADA
        #SI ESTÁ RESERVADA METER EL CÓDIGO EN IDS Y DARLE ENTER PARA QUE IMPRIMA LA OF
    #IMPRIMIR HOJA DE REGISTRO
    try:
        with db_conn_po('ODBC Driver 17 for SQL Server', 'MENITSISTEPLANT', 'marcadongvs', 'landerurgoiti', '2023Urria!') as cursor:
            # We truncate the table
            consulta = 'SELECT PartNumber, Fichero, TipoPieza, Marcadora, TipoMarcado FROM ProgramaMarcado WHERE PartNumber LIKE ?'
            cursor.execute(consulta, [KH])
            resultado = cursor.fetchone() #ESTA FUNCIÓN SACA LA OF, EL CÓDIGO ARTÍCULO Y EL SN DE FUNDIDO
            Fichero=resultado[1]
            ASSY=resultado[2]
            MARCADORA=resultado[3]
            Tipo_Marcado=str(resultado[4])
    except:
        print('No hay un programa definido ppara esta pieza')
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
    DM = Corchete + ' ' + RS + ' ' + Doce + ' ' + GS + ' ' + MFR_hex + ' ' + GS + ' ' + SER_hex + ' ' + GS + ' ' + PNR_hex + ' ' + RS + ' ' + EOT
    ASSY_hex = ' '.join([hex(ord(c))[2:].zfill(2) for c in ASSY])
   
    for i in range(0, 2):

        if MARCADORA=='VERTICAL':
            Marcar_Texto = '10 02 ' + LT + ' 00 36' + ' ' +Activar_Marcadora_Horizontal


        elif MARCADORA=='HORIZONTAL':
            Marcar_Texto = '10 02 ' + LT + ' 00 36'

        
        if Tipo_Marcado == '1':#tipo de marcado con DM único y rectangular

            trama = Marcar_Texto + ' ' + Punto + ' 0a ' \
                    + Data_Matrix_Rect_abre + ' ' + DM + ' ' + Data_Matrix_Rect_cierra + ' 0a ' \
                    + MFR_hex + ' 0a ' \
                    + SER_hex + ' 0a ' \
                    + PNR_hex + ' 0a ' \
                    + ASSY_hex

        elif Tipo_Marcado == '2':#tipo de marcado con DM único y cradrado

            trama = Marcar_Texto + ' ' + Punto + ' 0a ' \
                    + DM_SQ_abre + ' ' + DM + ' ' + DM_SQ_cierra + ' 0a ' \
                    + MFR_hex + ' 0a ' \
                    + SER_hex + ' 0a ' \
                    + PNR_hex + ' 0a ' \
                    + ASSY_hex
            
        elif Tipo_Marcado=='3':#tipo de marcado antiguo, con 2 DM cradrados
            DM1 = MFR_hex + ' ' + SER_hex
            DM2=PNR_hex
            trama = Marcar_Texto + ' ' + Punto + ' 0a ' \
                    + DM_SQ_abre + ' ' + DM1 + ' ' + DM_SQ_cierra + ' 0a ' \
                    + MFR_hex + ' 0a ' \
                    + SER_hex + ' 0a ' \
                    + DM_SQ_abre + ' ' + DM2 + ' ' + DM_SQ_cierra + ' 0a ' \
                    + PNR_hex + ' 0a ' \
                    + ASSY_hex

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
    if hex_string == ACK:
        resultado_label2.config(text='Texto transmitido e inicio de marcado...')
        print('Marcando...')
    elif hex_string == NACK:
        resultado_label2.config(text='Error en la trama')

    time.sleep(30)

    resultado_label2.config(text='Marcado terminado')
    client_socket.close()
    estado()
    #Registrar el marcado
    try:
        with db_conn_po('ODBC Driver 17 for SQL Server', 'MENITSISTEPLANT', 'marcadongvs', 'landerurgoiti', '2023Urria!') as cursor:
                    # We truncate the table
                    hora=datetime.now()
                    OFint=int(OF)
                    Registro = f'({KH}, {OFint}, {hora})'
                    print(Registro)
                    consulta = f'INSERT INTO RegistroMarcado (PartNumber, Orden, FechaMarcado) VALUES (?, ?, ?)'
                    cursor.execute(consulta, KH, OFint, hora)
                    cursor.commit()
                                    
                    #cursor.execute(consulta, [Registro])
    except po.Error as ex:
        # Capturar la excepción y mostrar información detallada sobre el error
        print("Se produjo un error al insertar los datos en la base de datos:")
        print("Tipo de error:", type(ex).__name__)
        print("Mensaje de error:", ex)
    cursor.close()
    codigo_textbox.insert(0, '')

    

# Crear la ventana
ventana = tk.Tk()
ventana.geometry('500x400+100+100')
ventana.title('Marcado NGV4')

# Botón de cerrar
boton_cerrar = tk.Button(ventana, text='Cerrar', width=10, command=ventana.destroy)
boton_cerrar.place(x=400, y=200)

# Botón RESET
boton_reset = tk.Button(ventana, text='Reset', width=10, command=lambda: reset_marcadora())
boton_reset.place(x=400, y=150)

# label 1
resultado_label1 = tk.Label(ventana, text='SER ->', font=('Arial', 20, 'bold'), justify='left')
resultado_label1.place(x=80, y=200, anchor=tk.CENTER)

# Crear la caja de texto
codigo_textbox = tk.Entry(ventana, font=('Arial', 20, 'bold'), justify='center')
codigo_textbox.place(x=250, y=200, width=250, height=50, anchor=tk.CENTER)

codigo_textbox.bind('<Return>', lambda event: marcar(codigo_textbox.get()))

# label 2
resultado_label2 = tk.Label(ventana, text='', font=('Arial', 10), justify='left')
resultado_label2.place(x=250, y=300, anchor=tk.CENTER)

# Cargar la imagen utilizando opencv
imagen1 = Image.open("M:\\OLANET\\G02-13 MITUTOYO CRYSTA APEX\\NGVLEAN\\scripts\\Logo_Egile_Mechanics.jpg")
imagen1 = imagen1.resize((128, 73))
photo = ImageTk.PhotoImage(imagen1)

# label 4
resultado_label_egile = tk.Label(ventana, image=photo)
resultado_label_egile.place(x=70, y=40, anchor=tk.CENTER)

estado()

ventana.mainloop()
