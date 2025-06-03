
import tkinter.messagebox as messagebox

def buscar_y_pegar():
    IDS = 0

IDS = 0
while IDS == 0:
    IDS=buscar_y_pegar()#Abre la OF y la Imprime, necesita el SER y el codigo de articulo para el logo de color
    if IDS==2:
        print ('Salida')