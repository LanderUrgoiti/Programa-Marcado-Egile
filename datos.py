import pandas as pd


PROGRAMA = pd.DataFrame({
    'PARTNUMBER': ['KH37848', 'KH37852', \
                    'KH37831', 'KH37833', \
                    'KH37815', 'KH37818', \
                    'KH19852', 'KH19853', \
                    'KH19863', 'KH19864', \
                    'KH19872', 'KH19873', \
                    'KH18990', 'KH18995', \
                    'KH88084', 'KH88086', \
                    'KH88091', 'KH88092'],

#TIEMPO 4X STD Y BORO SIN EL TOUCH PROBE 47, 56
                    
    'FICHERO': ['KH37848', 'KH37852', \
                    'KH37831', 'KH37833', \
                    'KH378151', 'KH37818', \
                    'KH19852', 'KH19853', \
                    'KH19863', 'KH19864', \
                    'KH198721', 'KH19873', \
                    'KH189901', 'KH18995', \
                    'KH88084', 'KH88086', \
                    'KH88091', 'KH88092'],
    'MARCADORA': ['VERTICAL', 'VERTICAL', \
                    'VERTICAL', 'VERTICAL', \
                    'VERTICAL', 'VERTICAL', \
                    'VERTICAL', 'VERTICAL', \
                    'VERTICAL', 'VERTICAL', \
                    'VERTICAL', 'VERTICAL', \
                    'HORIZONTAL', 'HORIZONTAL', \
                    'VERTICAL', 'VERTICAL', \
                    'VERTICAL', 'VERTICAL'],
    'TIPO MARCADO': ['2', '2', \
                    '2', '2', \
                    '2', '2', \
                    '2', '2', \
                    '2', '2', \
                    '1', '1', \
                    '1', '1', \
                    'KH88084', 'KH88086', \
                    'KH88091', 'KH88092'],

    'MFR': ['MFR', 'MFR', \
            'MFR', 'MFR', \
            'MFR', 'MFR', \
            'MFR', 'MFR', \
            'MFR', 'MFR', \
            'MFR', 'MFR', \
            'MFR', 'MFR', \
            'MFR', 'MFR', \
            'MFR', 'MFR'],

    'SER': ['SER', 'SER', \
            'SER', 'SER', \
            'SER', 'SER', \
            'SER', 'SER', \
            'SER', 'SER', \
            'SER', 'SER', \
            'SER', 'SER', \
            'SER', 'SER', \
            'SER', 'SER'],
    'PNR': ['PNR', 'PNR', \
            'PNR', 'PNR', \
            'PNR', 'PNR', \
            'PNR', 'PNR', \
            'PNR', 'PNR', \
            'PNR', 'PNR', \
            'PNR', 'PNR', \
            'PNR', 'PNR', \
            'PNR', 'PNR'],
    'ASSY': ['', '', \
            '', '', \
            '', '', \
            'ASSY', 'ASSY', \
            'ASSY', 'ASSY', \
            'ASSY', 'ASSY', \
            'ASSY', 'ASSY', \
            'ASSY', 'ASSY', \
            'ASSY', 'ASSY'],
    })