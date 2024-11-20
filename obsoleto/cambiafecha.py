import os
import datetime
import time

# Ruta al archivo
#archivo = r'\\srv5\\Calidad\\12. CCC\\HISTORICO\\1561\\CCC_1561.xlsx'
archivo = r'C:\\Users\\lander.urgoiti\\Desktop\\Lander Urgoiti\\Libro1.xlsx'

fecha_deseada = datetime.datetime(2023, 11, 27, 15, 22)

# Convertir fecha deseada a formato de timestamp
timestamp = time.mktime(fecha_deseada.timetuple())

# Cambiar la fecha de modificación del archivo
os.utime(archivo, (timestamp, timestamp))

print(f"La fecha de modificación del archivo {archivo} se ha cambiado a {fecha_deseada.strftime('%d/%m/%Y')}")