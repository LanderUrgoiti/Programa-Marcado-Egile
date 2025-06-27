import os
import sys
import json

kh = 'KH37848'

def get_data_path(filename):
    """Obtiene la ruta del archivo en desarrollo o compilado"""
    if hasattr(sys, '_MEIPASS'):
        # Si el script está congelado (ejecutable .exe)
        base_path = os.path.dirname(sys.executable)
    else:
        # En desarrollo (script .py)
        base_path = os.path.dirname(os.path.abspath(__file__))
        #base_path = os.path.join(base_path, '..', 'data')

    return os.path.join(base_path, filename)

# Ejemplo de uso
config_path = get_data_path('programas.json')
with open(config_path, 'r', encoding='utf-8') as f:
    config = json.load(f)
print(config['horizontal'])
hori = config['horizontal']
hori += 1
config['horizontal'] = hori
with open(config_path, 'w', encoding='utf-8') as f:
    json.dump(config, f, indent=4)

for item in config["programa"]:
    if item["partnumber"] == "KH18990":
        mfr_value = item["mfr"]
        print("MFR:", mfr_value)
        break
Fichero=
MARCADORA=resultado[2]
Tipo_Marcado=str(resultado[3])
MFR = resultado[4]
SER = resultado[5]
PNR = resultado[6]
ASSY = resultado[7]    
REV = resultado[8]