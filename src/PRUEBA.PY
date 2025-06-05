import datos

pn='KH37815'



PNR = datos.PROGRAMA.loc[datos.PROGRAMA['PARTNUMBER']==pn, 'PNR'].values[0]