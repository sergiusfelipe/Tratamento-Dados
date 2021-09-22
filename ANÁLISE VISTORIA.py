import tkinter as tk
from tkinter import filedialog,ttk
import pandas as pds
from math import sin, cos, sqrt, atan2, radians, asin
from datetime import datetime, timedelta,date,time
import random

def seconds_hours(duration):
    days, seconds = duration.days, duration.seconds
    hours = days * 24 + seconds // 3600
    minutes = (seconds % 3600) // 60
    seconds = (seconds % 60)
    olha_a_hora = str(hours) + ':' + str(minutes) + ':' + str(seconds)

    return olha_a_hora

def get1():
    pl1 = filedialog.askopenfilename()
    pl1_1 = pds.read_csv(pl1, error_bad_lines=False)
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


def excel1():
        csv1 = get1()
        csv = csv1[csv1['Timestamp'].notna()]
        try:
            csv = csv[csv['MultiChoiceSelection: TIPO/DT/CC/Madeira/EH/CASAS'] == 'DT' and csv['MultiChoiceSelection: TIPO/DT/CC/Madeira/EH/CASAS'] == 'EH']
        except:
            try:
                csv = csv[csv['MultiChoiceSelection: TIPO/DT/CC/Madeira/EH'] == 'DT' and csv['MultiChoiceSelection: TIPO/DT/CC/Madeira/EH/CASAS'] == 'EH']
            except:
                print("Num deu")
        csv = csv[csv['Timestamp'] != 'False']
        csv = csv[csv['Timestamp'] != '-1']
        csv = csv[csv['Timestamp'] != '9/150']
        csv = csv[csv['Timestamp'] != '9/300']
        csv = csv[csv['Timestamp'] != '8/300']
        csv = csv[csv['Timestamp'] != '11/500']
        csv = csv[csv['Timestamp'] != '4']
        csv = csv[csv['Timestamp'] != ' ']
        csv = csv[csv['Timestamp'] != '10/150']
        csv = csv[csv['Timestamp'] != '11/300']
        csv = csv[csv['Timestamp'] != '10.5/300']
        csv = csv[csv['Timestamp'] != '10/300']
        csv = csv[csv['Timestamp'] != '12/600']
        csv = csv[csv['Timestamp'] != '11/2000']
        csv = csv[csv['Timestamp'] != '10.5/600']
        csv = csv[csv['Timestamp'] != '10.5/150']
        csv = csv[csv['Timestamp'] != '12/1000']
        csv = csv[csv['Timestamp'] != '12/300']
        csv = csv[csv['Timestamp'] != '11/150']
        csv = csv[csv['Timestamp'] != '7.5/100']
        csv = csv[csv['Timestamp'] != '11/600']
        csv = csv[csv['Timestamp'] != '10.5/1000']
        csv = csv[csv['Timestamp'] != '8/100']
        csv = csv.reset_index(drop=True)
        imp = pds.DataFrame()
        segundos = []
        #print(csv)
        for k in range(0,len(csv.index)):
            tempo = str(csv.loc[k,'Timestamp'])
            a = tempo.split(' ')
            print(a)
            data = a[0]
            horas = a[1]
            data1 = data.split('-')
            horas1 = horas.split(':')
            hora_vis = timedelta(hours=int(horas1[0]),minutes=int(horas1[1]),seconds=int(horas1[2]))
            x1 = int(hora_vis.seconds)+(3600*24*int(data1[2]))+(3600*24*30*int(data1[1]))
            segundos.append(x1)
            
        csv['SEGUNDOS'] = segundos
        csv2 = csv.sort_values('SEGUNDOS',ascending=True)
        csv1 = csv2.reset_index(drop=True)
        #csv1.to_excel('BASE.xlsx')
        dia_vistoria = []
        primeiros_m = []
        primeiros_t = []
        ultimos_m = []
        ultimos_t = []
        media_tempo_m = []
        media_tempo_t = []
        postes_ma = []
        postes_ta = []
        
        hora_atua = 0
        hora_anterior = 0
        postes = 1
        med_m = 0
        med_t = 0
        prim_m = 0
        prim_t = 0
        ult_m = 0
        ult_t = 0
        dia_ant = 0
        postes_m = 0
        postes_t = 0
        dia_vistoria_ant = 0
        horario = 0
        primeira_dia = 0
        
        meio_dia = timedelta(hours=12,minutes=30)
        for i in range(0,len(csv.index)):
            tempo = csv1.loc[i,'Timestamp']
            a = tempo.split(' ')
            data3 = a[0]
            data2 = data3.split('-')
            d = str(data2[2]) + '/' + str(data2[1]) + '/' +str(data2[0])
            data = d
            horas = a[1]
            data1 = data.split('/')
            horas1 = horas.split(':')
            dia = int(data1[0])
            hora_vis = timedelta(hours=int(horas1[0]),minutes=int(horas1[1]),seconds=int(horas1[2]))
            if primeira_dia == 0:
                primeirissima = horas.split(':')
                #print('entrou')
                if hora_vis.seconds > meio_dia.seconds:
                    primeiros_m.append('FALTA')
                    ultimos_m.append('FALTA')
                primeira_dia =1
            if dia != dia_ant and dia_ant > 0:
                ult_t = horario
                u = ult_t.split(':')
                q = timedelta(hours=int(u[0]),minutes=int(u[1]),seconds=int(u[2]))
                h = timedelta(hours=int(primeirissima[0]),minutes=int(primeirissima[1]),seconds=int(primeirissima[2]))
                if q.seconds < meio_dia.seconds:
                    ultimos_t.append('FALTA')
                    primeiros_t.append('FALTA')
                else:
                    primeiros_t.append(prim_t)
                    ultimos_t.append(str(ult_t))
                #print(h.seconds)
                #print(meio_dia.seconds)
                if hora_vis.seconds > meio_dia.seconds:
                    primeiros_m.append('FALTA')
                    ultimos_m.append('FALTA')
                else:
                    if q.seconds < meio_dia.seconds:
                        #print(ult_m)
                        ultimos_m.append(str(horario))
                    else:
                        ultimos_m.append(str(ult_m))
                dia_vistoria.append(dia_vistoria_ant)
                postes_t = postes-postes_m-1
                postes_ma.append(postes_m)
                postes_ta.append(postes_t)
                postes = 1
            if hora_anterior == 0 or hora_vis.seconds < hora_anterior or (dia != dia_ant and hora_vis.seconds < meio_dia.seconds):
                if hora_vis.seconds <= meio_dia.seconds:
                    prim_m = horas
                    primeiros_m.append(str(prim_m))    
            
            if hora_vis.seconds >= meio_dia.seconds:
                #print(hora_anterior)
                #print(meio_dia.seconds)
                if hora_anterior <= meio_dia.seconds:
                    #print('entrou')
                    ult_m = horario
                    #print(ult_m)
                    postes_m = postes-1
                    prim_t = str(horas)
                
             
            postes=postes+1
            dia_vistoria_ant = data
            dia_ant = dia
            horario = horas
            hora_anterior = hora_vis.seconds
        
        dia_vistoria.append(data)
        
        primeirissima = horario.split(':')
        h = timedelta(hours=int(primeirissima[0]),minutes=int(primeirissima[1]),seconds=int(primeirissima[2]))
        d_i_a = len(dia_vistoria)
        #print(d_i_a)
        if h.seconds < meio_dia.seconds:
            ultimos_m.append(str(horario))
            ultimos_t.append('FALTA')
            primeiros_t.append('FALTA')
        else:
            if d_i_a == 1 and len(primeiros_m) == 0:
                ultimos_m.append('FALTA')
                primeiros_m.append('FALTA')
            else:
                ultimos_m.append(str(ult_m))
            ultimos_t.append(str(horario))
            primeiros_t.append(prim_t)
        
        
            
        postes_t = postes-postes_m-1
        postes_ma.append(postes_m)
        postes_ta.append(postes_t)
        
        vistoriador = []
        projeto = []
        segundos_m = []
        segundos_t = []
        media_vist = []
        total_postes = []
        total_segundos = []
        interv_m = []
        interv_t = []
        intervalo = []
        p1 = len(primeiros_m)
        p2 = len(primeiros_t)
        try:
            ultimos_m.remove('0')
        except:
            print("Sem 0")
        print(dia_vistoria)
        print(primeiros_m)
        print(ultimos_m)
        print(primeiros_t)
        print(ultimos_t)
        if p1 == 0:
            for m in range(0,len(dia_vistoria)):
                primeiros_m.append("00:00:00")
                ultimos_m.append("00:00:00")
        if p2 == 0:
            for n in range(0,len(dia_vistoria)):
                primeiros_t.append("00:00:00")
                ultimos_t.append("00:00:00")
        for j in range(0,len(dia_vistoria)):
            total_postes.append(postes_ma[j]+postes_ta[j])
            try:
                vistoriador.append(csv1.loc[i,'FreeText: COLABORADOR'])
            except:
                try:
                    vistoriador.append(csv1.loc[i,'FreeText: VISTORIADOR'])
                except:
                    try:
                        vistoriador.append(csv1.loc[i,'FreeText: TÉCNICO DE VISTORIA'])
                    except:
                        vistoriador.append('')
            projeto.append(csv1.loc[i,'Folder name'])
            #print(primeiros_m)
            horas2 = primeiros_m[j].split(':')
            horas3 = ultimos_m[j].split(':')
            try:
                hora_vis1 = timedelta(hours=int(horas2[0]),minutes=int(horas2[1]),seconds=int(horas2[2]))
                hora_vis2 = timedelta(hours=int(horas3[0]),minutes=int(horas3[1]),seconds=int(horas3[2]))
            except:
                hora_vis1 = timedelta()
                hora_vis2 = timedelta()
            interv_m.append(hora_vis2)
            delta_segundos_m = hora_vis2.seconds-hora_vis1.seconds
            segundos_m.append(delta_segundos_m)
            horas2 = primeiros_t[j].split(':')
            horas3 = ultimos_t[j].split(':')
            try:
                hora_vis1 = timedelta(hours=int(horas2[0]),minutes=int(horas2[1]),seconds=int(horas2[2]))
                hora_vis2 = timedelta(hours=int(horas3[0]),minutes=int(horas3[1]),seconds=int(horas3[2]))
            except:
                hora_vis1 = timedelta()
                hora_vis2 = timedelta()
            interv_t.append(hora_vis1)
            delta_segundos_t = hora_vis2.seconds-hora_vis1.seconds
            segundos_t.append(delta_segundos_t)
        
        for l in range(0,len(dia_vistoria)):
            if primeiros_m[l] == 'FALTA' or ultimos_m[l] == 'FALTA' or primeiros_t[l] == 'FALTA' or ultimos_t[l] == 'FALTA':
                intervalo.append('SEM INTERVALO')
            else:
                intervalo.append(str(interv_t[l]-interv_m[l]))
            total_segundos.append(str(timedelta(seconds=segundos_t[l]+segundos_m[l])))
            media_total = str(timedelta(seconds=segundos_t[l]+segundos_m[l])/total_postes[l])
            arredondar = media_total.split('.')
            media_vist.append(arredondar[0])
         
        imp['VISTORIADOR'] = vistoriador
        imp['PROJETO'] = projeto
        imp['DIA'] = dia_vistoria
        imp['PRIMEIRO MANHA'] = primeiros_m
        try:
            imp['ULTIMO MANHA'] = ultimos_m
        except:
            ultimos_m.pop(-1)
            imp['ULTIMO MANHA'] = ultimos_m
        #imp['MEDIA MANHA'] = media_tempo_m
        #imp['POSTES MANHA'] = postes_ma
        imp['PRIMEIRO TARDE'] = primeiros_t
        try:
            imp['ULTIMO TARDE'] = ultimos_t
        except:
            ultimos_t.pop(-1)
            imp['ULTIMO TARDE'] = ultimos_t
        imp['HORAS TRABALHADAS'] = total_segundos
        imp['INTERVALO'] = intervalo
        imp['MEDIA DIA'] = media_vist
        #imp['MEDIA TARDE'] = media_tempo_t
        #imp['POSTES TARDE'] = postes_ta
        imp['TOTAL POSTE'] = total_postes
        
        nome = str(projeto[0]) + '_' + str(date.today()) + str(random.randint(1,100000)) + '.xlsx'
        print(nome)
        try:
            imp.to_excel(nome)
        except:
            imp.to_excel('VISTORIA_' + str(date.today()) + str(random.randint(1,100000)) + '.xlsx')
        #concluido()
        
        
        
if __init__ == "__main__":

    root= tk.Tk()

    canvas1 = tk.Canvas(root, width = 300, height = 300, bg = 'black')
    canvas1.pack()  

    browseButton_Excel_1 = tk.Button(text='PROCESSAR VISTORIA', command=excel1, bg='white', fg='red', font=('helvetica', 12, 'bold'))
    canvas1.create_window(150, 150, window=browseButton_Excel_1)
    canvas1.create_text(125,295,fill="white",text="Desenvolvido por Sergio Tavora")

    root.mainloop()