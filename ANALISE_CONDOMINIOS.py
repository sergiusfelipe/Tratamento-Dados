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

def ajuste(palavra):

    spec_chars = ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','X','Z','W','Y','Á','À','Ã','Â','É','È','Ê','Í','Ì','Î','Ó','Ò','Õ','Ô','Ú','Ù','Û','Ç']
    
    o = []
    for letra in palavra:
        if letra in spec_chars:
            o.append(letra)
                #print(letra)
               
    p = str(''.join(o))

    return str(p)

def excel1():
        print('SELECIONAR A PLANILHA TIX')
        planilha1 = get1()
        planilha1.columns = planilha1.columns.str.strip()
        print("Carregamento 1 OK")

        print('SELECIONAR A PLANILHA MOB')
        planilha2 = get1()
        planilha2.columns = planilha2.columns.str.strip()
        print('Carregamento 2 OK')

        nomes_1 = []
        key_1 = []

        nomes_2 = []
        key_2 = []

        for i in range(0,len(planilha2['CONDOMINIO'])):
            nomes_1 = planilha2.loc[i,'CONDOMINIO'].split(' ')
            chave  = nomes_1[-1] + "_" + str(planilha2.loc[i,'Bairro'])
            key_1.append(chave)

        planilha2['KEY'] = key_1

        for i in range(0,len(planilha1['CONDOMINIO'])):
            nomes_2 = planilha1.loc[i,'CONDOMINIO'].split(' ')
            chave  = nomes_2[-1] + "_" + str(planilha1.loc[i,'Bairro'])
            key_2.append(chave)

        planilha1['KEY'] = key_2
        planilha2['KEY'] = key_1

        bairro = []
        etapa_1 = []
        etapa_2 = []
        etapa_3 = []
        cond1 = []
        cond2 = []
        cond3 = []
        count = []


        for i in range(0,len(planilha1['CONDOMINIO'])):
            nomes_2 = planilha1.loc[i,'CONDOMINIO'].split(' ')
            count.append(len(nomes_2))
            x = 'NOK'
            y = 'NOK'
            z = 'NOK'
            condominio1 = ''
            condominio2 = ''
            condominio3 = ''
            for j in range(0,len(planilha2['CONDOMINIO'])):
                nomes_1 = planilha2.loc[j,'CONDOMINIO'].split(' ')
                #print(nomes_2[-1] , ', ',nomes_1[-1])
                if True == True:
                    bairro1 = 'OK'
                    a = ajuste(nomes_2[-1])
                    b = ajuste(nomes_1[-1])
                    print(a, ', ',b)
                    if a == b:
                        etapa_11 = 'OK'
                        try:
                            a = ajuste(nomes_2[-2])
                            b = ajuste(nomes_1[-2])
                            print(a, ', ',b)
                            if a == b:
                                etapa_21 = 'OK'
                                try:
                                    a = ajuste(nomes_2[-3])
                                    b = ajuste(nomes_1[-3])
                                    print(a, ', ',b)
                                    if a == b:
                                        etapa_31 = 'OK'
                                    else:
                                        etapa_31 = 'NOK'
                                except:
                                    etapa_31 = 'NOK'
                            else:
                                etapa_21 = 'NOK'
                                etapa_31 = 'NOK'
                        except:
                            etapa_21 = 'NOK'
                            etapa_31 = 'NOK'
                    else:
                        etapa_11 = 'NOK1'
                        etapa_21 = 'NOK'
                        etapa_31 = 'NOK'
                else:
                    bairro1 = 'NOK'
                    etapa_11 = 'NOK'
                    etapa_21 = 'NOK'
                    etapa_31 = 'NOK'

                if etapa_11 == 'OK':
                    x = 'OK'
                    condominio1 = planilha2.loc[j,'CONDOMINIO']
                    if etapa_21 == 'OK':
                        y = 'OK'
                        condominio2 = planilha2.loc[j,'CONDOMINIO']
                        if etapa_31 == 'OK':
                            z = 'OK'
                            condominio3 = planilha2.loc[j,'CONDOMINIO']

            #bairro.append(bairro1)
            etapa_1.append(x)
            etapa_2.append(y)
            etapa_3.append(z)
            cond1.append(condominio1)
            cond2.append(condominio2)
            cond3.append(condominio3)
            


        #planilha1['LOC'] = bairro
        planilha1['ETAPA 1'] = etapa_1
        planilha1['COND_MOB_1'] = cond1
        planilha1['ETAPA 2'] = etapa_2
        planilha1['COND_MOB_2'] = cond2
        planilha1['ETAPA 3'] = etapa_3
        planilha1['COND_MOB_3'] = cond3
        planilha1['COUNT'] = count



        planilha1.to_excel(r'C:\Users\sergio.tavora\Desktop\DEMANDA CONDOMINIOS LUIS\TIX.xlsx')
        planilha2.to_excel(r'C:\Users\sergio.tavora\Desktop\DEMANDA CONDOMINIOS LUIS\MOB.xlsx')
                
        '''chave = []
        var1 = tk.DoubleVar()
        barra1 = ttk.Progressbar(root, variable = var1, maximum=len(planilha1.index),mode='determinate')
        barra1.pack()
        for i in range(0,len(planilha1.index)):
            palvra = str(planilha1.loc[i,'KEY'])
            cont = int(float(len(palvra))/2)
            firstpart, secondpart = palvra[:cont], palvra[cont:]
            if firstpart == secondpart:
                chave.append(firstpart)
            else:
                chave.append(str(planilha1.loc[i,'KEY']))
            #var1.set(i)
            #root.update()
        planilha1['KEY'] = chave
        cruzamento = pds.merge(planilha1,planilha2,on='KEY',how='left')
        print('Processamento 2 OK')
        info = []
        for j in range(0,len(cruzamento.index)):
            #ANALISE DE INFORMACAO
            if str(cruzamento.loc[j,'DescricaoStatusContrato']) == '' or str(cruzamento.loc[j,'DescricaoStatusContrato']) == 'nan' or str(cruzamento.loc[j,'DescricaoStatusContrato']) == str(cruzamento.loc[j,'IDContrato']):
                info.append("S/I")
            else:
                info.append("OK")
            
            
        cruzamento['INFO'] = info
        writer = pds.ExcelWriter('ANALISE.xlsx', engine='xlsxwriter')
        cruzamento.to_excel(writer, sheet_name='Sheet1', startrow=1, header=False, index=False)
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']
        (max_row, max_col) = cruzamento.shape
        column_settings = [{'header': column} for column in cruzamento.columns]
        worksheet.add_table(0, 0, max_row, max_col - 1, {'columns': column_settings})
        worksheet.set_column(0, max_col - 1, 12)
        writer.save()'''
        concluido()

if __name__ == "__main__":
 
    root= tk.Tk()

    canvas1 = tk.Canvas(root, width = 300, height = 300, bg = 'black')
    canvas1.pack()
    
    browseButton_Excel_1 = tk.Button(text='ANALISE', command=excel1, bg='white', fg='red', font=('helvetica', 12, 'bold'))
    canvas1.create_window(150, 150, window=browseButton_Excel_1)

    root.mainloop()