import tkinter as tk
from tkinter import filedialog,ttk
import pandas as pds
from geopy.geocoders import Bing, ArcGIS, Nominatim, Photon
from time import sleep
import simplekml
from math import sin, cos, sqrt, atan2, radians, asin
import utm
import time

#UTM
def distance_cartesian(x1, y1, x2, y2):
    dx = x1 - x2
    dy = y1 - y2

    return sqrt(dx * dx + dy * dy)

#DECIMAL
def dist(lat1,lon1,lat2,lon2):
    # approximate radius of earth in km
    try:
        lat1,lon1,lat2,lon2 = map(radians,[float(lat1),float(lon1),float(lat2),float(lon2)])
        R = 6371

        dlon = lon2 - lon1
        dlat = lat2 - lat1

        a = sin(dlat / 2)**2 + cos(lat1) * cos(lat2) * sin(dlon / 2)**2
        c = 2 * asin(sqrt(a))

        km = R * c
    except:
        km = 1000
        
    return km * 1000

def get1():
    pl1 = filedialog.askopenfilename()
    pl1_1 = pds.read_excel(pl1)
    pl1_2 = pds.DataFrame(pl1_1)
    
    return pl1_2

def end():
    return exit
    
def concluido():
    ok = tk.Toplevel(root)
    canvas2 = tk.Canvas(ok, width = 300, height = 300, bg = "#00ffff")
    botao = tk.Button(ok, text = 'Concluido', command = end, bg='white', fg='green', font=('helvetica', 12, 'bold'))
    canvas2.create_window(150, 150, window=botao)
    canvas2.pack()


def getPRE():
    
    pl1 = filedialog.askopenfilename()
    pl1_1 = pds.read_excel(pl1)
    data = pds.DataFrame(pl1_1)
    #data['NomeCidade'] = data['NomeCidade'].str.strip()
    caixas = get1()
    d = []
    ct = []
    distancia = []
    d_min = []
    d_ctoe = []
    d_min_1 = ""
    d_ctoe_1 = ""
    status = []
    classificacao = ""
   
    for i in range(0,len(data['SEQUENCIA'])): 
        start_time = time.time()
        for j in range(0, len(caixas['SEQUENCIA'])):
            la1 = data.loc[i,'LATITUDE']
            lo1 = data.loc[i,'LONGITUDE']
            la2 = caixas.loc[j,'LATITUDE']
            lo2 = caixas.loc[j,'LONGITUDE']
            l = dist(la1,lo1,la2,lo2)
            distancia.append(l)
            if distancia[j] <= min(distancia):
                perto = l
                caixa_px = caixas.loc[j,'SEQUENCIA']
            if distancia[j] <= 5:
                d_ctoe_1 = d_ctoe_1 + str(caixas.loc[j,'SEQUENCIA']) + ","
                d_min_1 = d_min_1 + str(l) + ","
                classificacao = "REPETIDO"
            
                         
        #print('CALCULANDO DISTANCIAS: ',i*100/len(data['HANDLE']),'%. ')
        #print(perto)
        #print(caixa_px)
        d.append(perto)
        ct.append(caixa_px)
        d_ctoe.append(d_ctoe_1[:-1])
        d_min.append(d_min_1[:-1])
        status.append(classificacao)
        distancia.clear()
        d_ctoe_1 = ""
        d_min_1 = ""
        classificacao = ""
        print('CALCULANDO DISTANCIAS: ',i*100/len(data['SEQUENCIA']),'%. ',"--- %s seconds ---" % (time.time() - start_time))
    
    data['DIST_MIN'] = d
    data['NXT_POSTE'] = ct
    data['POSTE_RAIO'] = d_ctoe
    data['DISTANCIAS_RAIO'] = d_min
    data['AVALIACAO'] = status
    
    data.to_excel("ANALISE_POSTES.xlsx")
    
    concluido()
    
def getDuplicada():
    
    data = get1()
    filtro = pds.DataFrame()
    postes = []
    for i in range(0,len(data['HANDLE'])):
        filtro = data[data['NXT_POSTE'] == data.loc[i,'NXT_POSTE']]
        filtro = filtro.reset_index(drop=True)
        poste = data.loc[i,'NXT_POSTE']
        j = 0
        count = len(filtro['HANDLE'])
        #print(count)
        lista = str(data.loc[i,'POSTE_RAIO']).split(',')
        if count > 1:
            for j in range(0,len(lista)):
            
                poste_lista = lista[j]
                print(poste_lista)
                filtro = data[data['NXT_POSTE'] == poste_lista]
                filtro = filtro.reset_index(drop=True)
                count_1 = len(filtro['HANDLE'])
                #print(count)
                #print(filtro)
                if count_1 < 1:
                #print(filtro)
                    poste = poste_lista
            
                #print("entrou " + str(j))
            
        data.at[i,'NXT_POSTE'] = poste
        postes.append(poste)    
        
    data['NXT_POSTE'] = postes
    data.to_excel("ANALISE_POSTE_1.xlsx")
    concluido()
    
if __name__ == "__main__":
    root= tk.Tk()

    canvas1 = tk.Canvas(root, width = 300, height = 300, bg = 'black')
    canvas1.pack()

    browseButton_Excel_1 = tk.Button(text='CRUZAR POSTES', command=getPRE, bg='white', fg='red', font=('helvetica', 12, 'bold'))
    browseButton_Excel_2 = tk.Button(text='DUPLICATAS', command=getDuplicada, bg='white', fg='red', font=('helvetica', 12, 'bold'))
    canvas1.create_window(150, 100, window=browseButton_Excel_1)
    canvas1.create_window(150, 200, window=browseButton_Excel_2)
    canvas1.create_text(125,295,fill="white",text="Desenvolvido por Sergio Tavora")

    root.mainloop()