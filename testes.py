from enum import _auto_null
from cv2 import COLOR_BGR2GRAY
import numpy as np
import cv2 as cv
import matplotlib.pyplot as plt
import openpyxl
import time

#Abre o Arquivo de Excel onde serão salvos os dados
wb = openpyxl.load_workbook('teste.xlsx')
ws = wb.active #Abre a aba ativa do arquivo de excel

#Função para associar um numero de iteração a uma coluna do excel
def find_letter(iterador):
    if iterador == 0:
        return 'A'
    if iterador == 1:
        return 'B'
    if iterador == 2:
        return 'C'
    if iterador == 3:
        return 'D'
    if iterador == 4:
        return 'E'
    if iterador == 5:
        return 'F'
    if iterador == 6:
        return 'G'
    if iterador == 7:
        return 'H'
    if iterador == 8:
        return 'I'
    if iterador == 9:
        return 'J'
    if iterador == 10:
        return 'K'
    if iterador == 11:
        return 'L'
    if iterador == 12:
        return 'M'
    if iterador == 13:
        return 'N'
    if iterador == 14:
        return 'O'
    if iterador == 15:
        return 'P'
    if iterador == 16:
        return 'Q'
    if iterador == 17:
        return 'R'
    if iterador == 18:
        return 'S'
    if iterador == 19:
        return 'T'
    if iterador == 20:
        return 'U'
    if iterador == 21:
        return 'V'
    if iterador == 22:
        return 'W'
    if iterador == 23:
        return 'X'
    if iterador == 24:
        return 'Y'
    if iterador == 25:
        return 'Z'
    if iterador == 26:
        return 'AA'
    if iterador == 27:
        return 'AB'
    if iterador == 28:
        return 'AC'
    if iterador == 29:
        return 'AD'
    if iterador == 30:
        return 'AE'
    if iterador == 31:
        return 'AF'
    if iterador == 32:
        return 'AG'
    if iterador == 33:
        return 'AH'
    if iterador == 34:
        return 'AI'
    if iterador == 35:
        return 'AJ'
    if iterador == 36:
        return 'AK'
    if iterador == 37:
        return 'AL'
    if iterador == 38:
        return 'AM'
    if iterador == 39:
        return 'AN'
    if iterador == 40:
        return 'AO'
    if iterador == 41:
        return 'AP'
    if iterador == 42:
        return 'AQ'
    if iterador == 43:
        return 'AR'
    if iterador == 44:
        return 'AS'
    if iterador == 45:
        return 'AT'
    if iterador == 46:
        return 'AU'
    if iterador == 47:
        return 'AV'
    if iterador == 48:
        return 'AW'
    if iterador == 49:
        return 'AX'
    if iterador == 50:
        return 'AY'

#Função para apagar todas as linhas do excel
def delete_all_rows():
    while ws['A1'].value != None:
        ws.delete_rows(1)
    return
    
#Função para achar a borda mais proxima do centro da imagem
def find_nearest_white(canny, target):
    nonzero = cv.findNonZero(canny)
    distances = np.sqrt((nonzero[:,:,0] - target[0]) ** 2 + (nonzero[:,:,1] - target[1]) ** 2)
    nearest_index = np.argmin(distances)
    return nonzero[nearest_index]

#Obtem o nome do arquivo de imagem
img = cv.VideoCapture('3.mp4')

#Define o tamanho da imagem
frame_width = int(img.get(3))
frame_height = int(img.get(4))
size = (frame_width, frame_height)

#Iniciador onde será salvo o video
result = cv.VideoWriter('00.mp4', cv.VideoWriter_fourcc(*'mp4v'), 10, size)
                 
#Input para receber a quantidade de linhas que aparecerão em cada folheto
number_of_angles = int(input("Digite o numero de linhas para cada folheto: "))

#Inicializ os dados do Excel vazios
delete_all_rows()

#Variavel para auxiliar na manipulação de linhas e colunas do excel
count = 2

#Variavel de linhas totais sendo igual ao numero de angulos multiplicado pelo nuemro de folhetos
linhas = number_of_angles*3

#Variavel de Angulo
angulo = 360/linhas

#Parte que define quantas colunas serão necessarias para o excel
iterador = 1
for iterador in range(linhas+1):
    letra = find_letter(iterador)
    ws[f'{letra}1'] = angulo*iterador

#Loop principal de iteracao
while True:
    ws[f'A{count}'] = str(count-1) 
    ret, frame = img.read()
    if not ret:
        break
    grey = cv.cvtColor(frame, cv.COLOR_BGRA2GRAY)
    kernel = np.ones((5,5), np.uint8)
    grey = cv.GaussianBlur(grey, (9,9), 0)
    grey = cv.morphologyEx(grey, cv.MORPH_OPEN, kernel)
    grey = cv.morphologyEx(grey, cv.MORPH_CLOSE, kernel)
    
    canny = cv.Canny(grey, 100, 200) 
    
    #centro da valvula
    TARGET = (int(canny.shape[1]/2), int(canny.shape[0]/2))
    
    
    # Ellipse parameters
    radius = 255
    axes = (radius, radius)
    angle = 0
    
    thickness = -1
    w0 = frame.shape[0]
    h0 = frame.shape[1]
    
    back = cv.imread('white.png')
    print(w0, h0)
    back = cv.resize(back, (h0, w0))
    frame = cv.subtract(back, frame)
    
    W = (255, 255, 255) #Branco
    R = 0, 0, 255 #Vermelho
    G = 0, 255, 0 #Verde
    B = 255, 0, 0 #Azul
    
    
    #alteração
    
    loop = 0
    i = 1
    k = 1
    
    iterador = 1
    while loop < linhas:
        letra = find_letter(iterador)
        if cv.waitKey(1) & 0xFF == ord('q'):
            img.release()
            result.release()
            cv.destroyAllWindows()
        mask = np.ones((w0,h0),dtype=np.uint8)
        src1_mask=cv.cvtColor(mask,cv.COLOR_GRAY2BGR)#change mask to a 3 channel image
        startAngle = (k*angulo)-1
        endAngle = (i*angulo)
        if loop == linhas-1:
            endAngle = 360
        cv.ellipse(src1_mask, TARGET, axes, angle, startAngle, endAngle, (255,255,255), thickness)
        cv.circle(src1_mask, TARGET, 5, (0,0,0), -1)
        mask_out=cv.subtract(src1_mask, frame)
        mask_out=cv.subtract(src1_mask,mask_out)
        grey = cv.GaussianBlur(mask_out, (9,9), 0)
        canny = cv.Canny(grey, 100, 200)
        try:
            pointAB = find_nearest_white(canny, TARGET)
        except:
            print(f"nenhum ponto branco em {startAngle} e {endAngle}")
        if startAngle >= 0 and endAngle <= 119:
            color = B
        if startAngle >= 120 and endAngle <= 239:
            color = G
        if startAngle >= 240 and endAngle <= 360:
            color = R
        cv.line(frame, (int(canny.shape[1]/2), int(canny.shape[0]/2)), (int(pointAB[0][0]), int(pointAB[0][1])), (color), 1)
        raio = np.sqrt((pointAB[0][0] - TARGET[0]) ** 2 + (pointAB[0][1] - TARGET[1]) ** 2)
        print(startAngle, endAngle, f'raio = {raio}')
        ws[f'{letra}{count}'] = raio
        iterador += 1
        i += 1
        k += 1
        loop += 1
        
    # cv.line(frame, (int(canny.shape[1]/2), int(canny.shape[0]/2)), (619, 100), (W), 1)
    # cv.line(frame, (int(canny.shape[1]/2), int(canny.shape[0]/2)), (1002, 319), (W), 1)
    # cv.line(frame, (int(canny.shape[1]/2), int(canny.shape[0]/2)), (618, 541), (W), 1)
     
    #Circulo externo
    cv.circle(frame, (int(canny.shape[1]/2), int(canny.shape[0]/2)), 255, (0, 255, 0), 2)
    
    #Pontos entre folhetos
    cv.circle(frame, (619, 100), 0, (0, 0, 255), 8)
    cv.circle(frame, (1002, 319), 0, (0, 0, 255), 8)
    cv.circle(frame, (618, 541), 0, (0, 0, 255), 8)
    
    count += 1
    cv.imshow('mask', frame)   
    result.write(frame)
    wb.save('teste.xlsx')
    #takes a screenshot
    if cv.waitKey(1) == ord('p'):
        cv.imwrite('screenshot.png', frame)

    if cv.waitKey(1) & 0xFF == ord('q'):
        break
    
img.release()
result.release()
cv.destroyAllWindows()