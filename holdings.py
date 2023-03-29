# noinspection PyUnresolvedReferences
from openpyxl import Workbook

# noinspection PyUnresolvedReferences
import ctypes

# noinspection PyUnresolvedReferences
from tkinter import *

# noinspection PyUnresolvedReferences
from tkinter import ttk

# noinspection PyUnresolvedReferences
import openpyxl

# noinspection PyUnresolvedReferences
import arrow

# noinspection PyUnresolvedReferences
import numpy as np

# noinspection PyUnresolvedReferences
from auxiliar import auxiliar

# noinspection PyUnresolvedReferences
import pandas

# noinspection PyUnresolvedReferences
import xlrd

# noinspection PyUnresolvedReferences
import xlwings as xw

# noinspection PyUnresolvedReferences
import time

# noinspection PyUnresolvedReferences
from openpyxl import load_workbook

# noinspection PyUnresolvedReferences
import xlsxwriter

# noinspection PyUnresolvedReferences
from PIL import ImageTk, Image

"""-------COMPLEMENTOS------"""

matriz0 = [
    [3900, "FUNDACION",0],
    [3300, "MANUTO",0],
    [3500, "LIBOFE",0],
    [3600, "JET MATCH",0],
    [3700, "MELL",0],
    [3800, "MICROCENTRO",0],
    [3400, "SAINT GALL",0],
    [8200, "SMG CORPORATE",0],
    [8000, "SMG INVESTMENT",0],
    [6500, "SMG SERVICES",0],
    [3000, "SMG SERVICIOS",0],
    [8100, "SWISS INVERSIONES",0],
    [7600, "INTERNACIONAL",0]
         ]


def agregar_0(month):
    if(month<10):
        return "0"+str(month)
    else:
        return month

def get_T1():
    if str(int(entry_m.get())-1)=="00":
        return ["12",str(int(entry_y.get())-1)]
    else:
        return [agregar_0(int(entry_m.get())-1),str(entry_y.get())]



"-------------------EMPIEZA LA INTERFAZ----------------"

#RAIZ
raiz=Tk()
raiz.title("SMG GROUP // " + arrow.now().format('DD-MM-YYYY'))
raiz.geometry("350x550")

#FONDO
img = ImageTk.PhotoImage(Image.open("SwissMedical.png"))
panel = Label(raiz, image = img)
panel.place(x=0, y=0, relwidth=1, relheight=1)

#SLIDE
# combo= ttk.Combobox()
# combo.place(x=65, y=50)
# combo["state"]="readonly"
# combo["values"]=["01","02","03","04","05","06","07","08","09","10","11","12"]

#CHECKBUTTONS
x=150
y=100
for i in range(len(matriz0)):
    matriz0[i][2]=Variable()
    checkbox = ttk.Checkbutton(text=matriz0[i][1], variable=matriz0[i][2], onvalue="on", offvalue="apagado")
    checkbox.place(x=30, y=y)
    y+=30

def filtro_flqls():

    #SEGREGADOR DE FLQLS SEGUN SOCIEDAD---------------------------------
    dir = "C:/Users/rodriaguirre/OneDrive - Swiss Medical S.A/Documents/General/02. Reportes Corporativos/01. Informe Financiero/FLQLS Holdings/"+entry_y.get()+"/"+entry_m.get()+"/flqls.xlsx"
    print("pase")

    flqls = pandas.read_excel(dir)

    writer = pandas.ExcelWriter("C:/Users/rodriaguirre/OneDrive - Swiss Medical S.A/Documents/General/02. Reportes Corporativos/01. Informe Financiero/FLQLS Holdings/"+entry_y.get()+"/"+entry_m.get()+"/flqls_filtradas.xlsx", engine='xlsxwriter')

    for i in range(13):
        flqls_filtrada = flqls.loc[flqls["Soc."] == matriz0[i][0]]
        flqls_filtrada.to_excel(writer, sheet_name=matriz0[i][1], index=False, header=False)
    writer.save()

    # SEGREGADOR DE BALANCES SEGUN SOCIEDAD------------------------------
    dir_blcs = r"C:/Users/rodriaguirre/OneDrive - Swiss Medical S.A/Documents/General/02. Reportes Corporativos/01. Informe Financiero/FLQLS Holdings/" + entry_y.get() + "/" + entry_m.get() + "/blcs.xlsx"

    blcs = pandas.read_excel(dir_blcs)

    #ETL BASE BALANCES-------
    blcs = blcs.iloc[:, 2:16]
    blcs.rename(columns={"Unnamed: 2": "Soc", "Unnamed: 4": "Mayor", "Unnamed: 14": "Variacion", "Unnamed: 7": "Descripcion","Unnamed: 10": "T+0", "Unnamed: 12": "T-1"}, inplace=True)
    blcs = blcs.drop(["Unnamed: 15"], axis=1)
    blcs.dropna(inplace=True, subset=['Soc'])
    blcs["data_type"] = blcs["Soc"].apply(lambda x: "STR" if isinstance(x, str) else "num")
    blcs = blcs.loc[blcs["data_type"] == "num"]
    # ETL BASE BALANCES-------

    writer = pandas.ExcelWriter(r"C:/Users/rodriaguirre/OneDrive - Swiss Medical S.A/Documents/General/02. Reportes Corporativos/01. Informe Financiero/FLQLS Holdings/"+ "2022" +"/"+ entry_m.get()+"/"+"blcs_filtradas.xlsx",engine='xlsxwriter')

    for i in range(13):
        # FILTRADOR DE EXTRAC
        blcs_filtrada = blcs.loc[blcs["Soc"] == matriz0[i][0]]

        # Write each dataframe to a different worksheet.
        blcs_filtrada.to_excel(writer, sheet_name=matriz0[i][1], index=False, header=False)

    # Close the Pandas Excel writer and output the Excel file.
    writer.save()

    ctypes.windll.user32.MessageBoxW(0, "Proceso termiando", "Confirmación", 0)

def armador_informes():
    matriz=[x for x in range(len(matriz0)) if matriz0[x][2].get() == "on"]

    flqls_filtradas="C:/Users/rodriaguirre/OneDrive - Swiss Medical S.A/Documents/General/02. Reportes Corporativos/01. Informe Financiero/FLQLS Holdings/"+str(entry_y.get())+"/"+str(entry_m.get())+"/flqls_filtradas.xlsx"
    flqls=xw.Book(flqls_filtradas)
    blcs_filtradas="C:/Users/rodriaguirre/OneDrive - Swiss Medical S.A/Documents/General/02. Reportes Corporativos/01. Informe Financiero/FLQLS Holdings/"+str(entry_y.get())+"/"+str(entry_m.get())+"/blcs_filtradas.xlsx"
    blcs=xw.Book(blcs_filtradas)
    print(str(agregar_0(int(entry_m.get())-1)))

    for i in range(len(matriz)):
        clave=int(matriz[i])
        dir="C:/Users/rodriaguirre/OneDrive - Swiss Medical S.A/Documents/General/02. Reportes Corporativos/01. Informe Financiero/"+ get_T1()[1]+"/"+ str(matriz0[matriz[i]][0])+" "+str(matriz0[matriz[i]][1])+ "/INFFIN Mensual "+ get_T1()[0] +" "+ get_T1()[1] +" "+matriz0[matriz[i]][1]+".xlsx"
        print("T-1: "+dir)
        informe=xw.Book(dir)
        time.sleep(6)
        dir2="C:/Users/rodriaguirre/OneDrive - Swiss Medical S.A/Documents/General/02. Reportes Corporativos/01. Informe Financiero/" + str(entry_y.get()) +"/"+ str(matriz0[matriz[i]][0])+" "+str(matriz0[matriz[i]][1]) +"/INFFIN Mensual "+ str(entry_m.get()) +" "+ str(entry_y.get()) +" "+matriz0[matriz[i]][1]+".xlsx"
        informe.save(dir2)
        print("T0: " + dir2)
        time.sleep(6)

        informe.sheets['Comprobación pre cierre'].range('B6').value=informe.sheets['resfi'].range('D104').value

        informe.sheets['Comprobación pre cierre'].range('B4').value= str(entry_m.get())+"/1/"+str(entry_y.get())

        informe.sheets['Detalle resfi'].range('A6:I300').value =None

        copia=flqls.sheets[matriz0[clave][1]].range("A1:I1000").value #COPIO PEGO FLQLS
        informe.sheets['Detalle resfi'].range('A6:I1000').value=copia

        copia=blcs.sheets[matriz0[clave][1]].range("A1:M1000").value #COPIO PEGO BLCS
        informe.sheets['Balance PyG'].range('A1:R1000').value = None
        informe.sheets['Balance PyG'].range('C7').value = str(matriz0[matriz[i]][0])+"-"+str(matriz0[matriz[i]][1])
        informe.sheets['Balance PyG'].range('K7').value = str(entry_m.get())+"/"+str(entry_y.get())
        informe.sheets['Balance PyG'].range('M7').value = str(get_T1()[0])+"/"+str(get_T1()[1])
        informe.sheets['Balance PyG'].range('O7').value = "Desviacion"
        informe.sheets['Balance PyG'].range('C9:O1000').value=copia

        informe.close()

    ctypes.windll.user32.MessageBoxW(0, "Proceso termiando", "Confirmación", 0)

matrizX=[x for x in range(len(matriz0)) if matriz0[x][2].get() == "on"]

def valor():
    ctypes.windll.user32.MessageBoxW(0, matrizX, "Confirmación", 0)

#BOTONOS Y LABELS
text = Label(panel, text="ARMADO INFORMES HOLDING")
text.place(x=30, y=5)

text2 = Label(panel, text="Mes:")
text2.place(x=30, y=50)

entry_m = Entry(raiz)
entry_m.place(x=65, y=52, width=20)


text2 = Label(panel, text="Año:")
text2.place(x=95, y=50)

entry_y = Entry(raiz)
entry_y.place(x=130, y=52, width=40)

botonarchivotransfer=Button(panel, text="Aceptar", height=1, width= 24,command=lambda:armador_informes())
botonarchivotransfer.place(x=75, y=500)

botonarchivotransfer2=Button(panel,text="Filtrar", height=1, width= 5,command=lambda:filtro_flqls())
botonarchivotransfer2.place(x=225, y=45)

raiz.mainloop()


"str(agregar_0(int(arrow.now().format('MM'))-2))"