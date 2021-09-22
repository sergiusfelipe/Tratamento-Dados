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


def excel1():
        print('SELECIONAR A PLANILHA DE FIBRAS')
        planilha1 = get1()
        #print(len(planilha1.index))
        print("Carregamento OK")
        '''planilha = pds.DataFrame(planilha1,columns=['NOME_LOCAL','STATUS_LOCAL','POINT_ID','PONTO_STATUS','ID_CLIENTE_ATRIX','ID_CONTRATO_ATRIX','ID_CLIENTE_ADAPTER','ID_CONTRATO_ADAPTER','PROJETO'])
        planilha.columns = planilha.columns.str.strip()
        planilha['KEY'] = planilha['ID_CLIENTE_ATRIX'].map(str) + planilha['ID_CONTRATO_ATRIX'].map(str) + planilha['ID_CLIENTE_ADAPTER'].map(str) + planilha['ID_CONTRATO_ADAPTER'].map(str)
        planilha = planilha[planilha['KEY'].notna()]
        planilha.columns = planilha.columns.str.strip()
        planilha = planilha[planilha['KEY'] != 'nannan']
        planilha = planilha[planilha['KEY'] != 'nannannan']
        planilha = planilha[planilha['KEY'] != 'nannannannan']
        planilha['KEY'] = planilha['KEY'].str.replace('nan','')
        planilha['KEY'] = planilha['KEY'].str.replace(' ','')
        planilha = planilha[planilha['KEY'] != ' ']
        planilha = planilha[planilha['KEY'] != '  ']
        planilha = planilha[planilha['KEY'] != '   ']
        planilha = planilha[planilha['KEY'] != '    ']
        planilha.columns = planilha.columns.str.strip()
        planilha = planilha[planilha['KEY'] != '']
        planilha = planilha.reset_index(drop=True)'''
        #print(planilha1)
        #print(len(planilha1.index))
        print("Processamento 1 OK")
        print('SELECIONAR BASE BI')
        planilha3 = get1()
        planilha2 = pds.DataFrame(planilha3,columns=['IDCliente','IDContrato','DescricaoStatusContrato','UFCidadeInstalacao','NomeCidadeInstalacao'])
        planilha2.columns = planilha2.columns.str.strip()
        planilha2['KEY'] = planilha2['IDCliente'].map(str) + planilha2['IDContrato'].map(str)
        print('Carregamento OK')
        chave = []
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
        writer.save()
        concluido()
 
if __name__ == "__main__":
 
    root= tk.Tk()

    canvas1 = tk.Canvas(root, width = 300, height = 300, bg = 'black')
    canvas1.pack()
    
    browseButton_Excel_1 = tk.Button(text='ANALISE', command=excel1, bg='white', fg='red', font=('helvetica', 12, 'bold'))
    canvas1.create_window(150, 150, window=browseButton_Excel_1)

    root.mainloop()