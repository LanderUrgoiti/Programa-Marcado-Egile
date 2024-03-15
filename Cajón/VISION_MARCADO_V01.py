import cv2
import numpy as np



'''
    dp: Resolución inversa de la acumuladora en relación con la resolución de la imagen. Generalmente, se establece en 1.
    minDist: La distancia mínima entre los centros de los círculos detectados. Aumentar este valor reducirá el número de círculos detectados.
    param1: Umbral de detección de bordes interno del algoritmo de detección de círculos de Hough. Un valor más bajo lo hará más sensible a los bordes.
    param2: Umbral de detección de centros del algoritmo de detección de círculos de Hough. Un valor más bajo lo hará más estricto en la detección de círculos.
    minRadius y maxRadius: El rango mínimo y máximo de los radios de los círculos a detectar.
'''
  
# Read image.
imagen = cv2.imread(r"\\srv5\Oficina Tecnica\03-AERO ENGINES\03-PROCESO\01-MARCADO\Medidas Control\DM1_20_02_2023.png", 0)

px=0.014427 #mm/pixel

# Aplicar un desenfoque para reducir el ruido
imagen_blur = cv2.GaussianBlur(imagen, (5, 5), 0)

# Aplicar la detección de círculos utilizando el algoritmo de Hough
circulos = cv2.HoughCircles(
    imagen_blur,
    cv2.HOUGH_GRADIENT,
    dp=1,
    minDist=0.05/px,
    param1=50,
    param2=10,
    minRadius=int(0.01/px),
    maxRadius=int(0.3/px)
)

# Verificar si se encontraron círculos
if circulos is not None:
    # Redondear y convertir las coordenadas y el radio a enteros
    circulos = np.round(circulos[0, :]).astype("int")

    # Dibujar los círculos encontrados
    for (x, y, r) in circulos:
        cv2.circle(imagen, (x, y), r, (0, 255, 0), 2)

    # Mostrar la imagen con los círculos detectados
    cv2.imshow("CIRCULOS DETECTADOS", imagen)
    cv2.waitKey(0)
    cv2.destroyAllWindows()
else:
    print("No se encontraron círculos en la imagen.")
