import tkinter as tk
from tkinter import filedialog,ttk
import pandas as pds

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

def planilha():
    
    data = get1()
    data['LOCAL_A'] = data['LOCAL_A'].str.strip()
    data['LOCAL_Z'] = data['LOCAL_Z'].str.strip()
    sep = []
    cd = []
    caixas = []
    emenda = []
    alvo = []
    
    cdoe = 0
    
    for i in range(0,len(data['NOME'])):
        sep.append(data.loc[i,'NOME'].split('-')[0])
        cd.append(data.loc[i,'LOCAL_A'].split('-')[0])
        
    data['TIPO'] = sep
    data['TIPO_C'] = cd
    
    for l in range(0,len(data['NOME'])):
        if data.loc[l,'TIPO_C'] == 'CTOE' or data.loc[l,'TIPO_C'] == 'CDOE':
            alvo.append('SIM')
        else:
            alvo.append('NAO')
      
    data['ALVO'] = alvo
    data.to_excel('resultado.xlsx')
    data1 = data[data['ALVO'] == 'SIM']
    data1 = data1.reset_index(drop=True)
    print(data1)
    x = 1
    emenda_1 = []
    for j in range(0,len(data1['NOME'])):
        caixa = data1.loc[j,'LOCAL_Z']
        caixas.append(caixa)
        emenda.append(cdoe)
        tipo = data1.loc[j,'TIPO_C']
        print('Linha ' + str(cdoe) + '. ' + str(caixa))
        while tipo != 'CDOE':
            try:
                #print(caixa)
                filtro = data1[data1['LOCAL_Z'] == caixa]
                filtro = filtro.reset_index(drop=True)
                #print(filtro)
                anterior = filtro.loc[0,'LOCAL_A']
                #print(anterior)
                caixas.append(anterior)
                emenda.append(cdoe)
                caixa = anterior
                tipo = filtro.loc[0,'TIPO_C']
                x = x + 1
            except:
                tipo = 'CDOE'
            
        for k in range(0,x):
            emenda_1.append(caixa)
        #emenda = [str(anterior) if k == str(cdoe)]
        cdoe = cdoe + 1
        x = 1
    
    data2 = pds.DataFrame()
    data2['CAIXA'] = caixas
    data2['CDOE'] = emenda_1
    print(data2)
    
    data2.to_excel('resultado_1.xlsx')
   
    concluido()

if __name__ == "__main__":
    root= tk.Tk()

    canvas1 = tk.Canvas(root, width = 300, height = 300, bg = 'black')
    canvas1.pack()

    browseButton_Excel_1 = tk.Button(text='MAPEAR', command=planilha, bg='white', fg='red', font=('helvetica', 12, 'bold'))
    canvas1.create_window(150, 150, window=browseButton_Excel_1)
    canvas1.create_text(125,295,fill="white",text="Desenvolvido por Sergio Tavora")

    root.mainloop()