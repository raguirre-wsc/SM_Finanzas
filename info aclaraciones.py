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
import tkinter

# noinspection PyUnresolvedReferences
from tkinter import ttk

# noinspection PyUnresolvedReferences
from PIL import ImageTk, Image

# noinspection PyUnresolvedReferences
from main import matriz_acl

pandas.set_option('display.max_rows', 1000)
pandas.set_option('display.max_columns', 2000)
pandas.set_option('display.width', 3000)
pandas.set_option('display.float_format', '{:.1f}'.format)

cias=[
    [3900, "FUNDACION",0],
    [3300, "MANUTO",0],
    [3500, "LIBOFE",0],
    [3700, "MELL",0],
    [3800, "MICROCENTRO",0],
    [3400, "SAINT GALL",0],
    [8200, "SMG CORPORATE",0],
    [8000, "SMG INVESTMENT",0],
    [6500, "SMG SERVICES",0],
    [3000, "SMG SERVICIOS",0],
    [8100, "SWISS INVERSIONES",0],
    [7600, "INTERNACIONAL",0],
    [1000, "SWISS MEDICAL",0],
    [1100, "ECCO",0],
    [7400, "GENERALES",0],
    [7300, "ART",0],
    [7200, "VIDA",0],
    [7000, "RETIRO",0]
         ]

lencias = len(cias)

dates = "C:/Users/rodrigo/Swiss Medical S.A/Finanzas UC - General/03. Posicion Financiera Diaria/01. Saldia/02. Seguros/" + arrow.now().format('YYYY') + "/" + arrow.now().format('MM.YYYY') + "/" + arrow.now().format('YYYY_MM_DD') + " SMG.xlsx"
df=pandas.read_excel(dates, sheet_name="Base")

df.columns = df.iloc[0]
df = df[1:]

df['Saldo Apertura Nuevo']=df['Saldo Apertura']

df.round({'Saldo Apertura Nuevo':-2})


df2=df.filter(["Soc.","Razón Social","Banco","Cuenta (Formato Interbanking)","Saldo Apertura","Transferencias","Creditos"])

df2 = df2.loc[df2['Banco'] == "ICBC"]

df2['Transferencias'] = pandas.to_numeric(df2['Transferencias'], errors='coerce')
df2['Creditos'] = pandas.to_numeric(df2['Creditos'], errors='coerce')

df2['Transferencias'] = df2['Transferencias'].fillna(0)
df2['Creditos']= df2['Creditos'].fillna(0)

df2["Saldo Apertura"] = df2["Saldo Apertura"].astype('int64')
df2["Transferencias"] = df2["Transferencias"].astype('int64')
df2["Creditos"] = df2["Creditos"].astype('int64')

df2['Saldo Apertura'] = df2['Saldo Apertura'].apply(lambda x: x/1000000)
df2['Transferencias'] = df2['Transferencias'].apply(lambda x: x/1000000)
df2['Creditos'] = df2['Creditos'].apply(lambda x: x/1000000)

df2["Saldo Apertura"] = df2["Saldo Apertura"].astype('float').round(1)
df2["Transferencias"] = df2["Transferencias"].astype('float').round(1)
df2["Creditos"] = df2["Creditos"].astype('float').round(1)

df3=df2

df3=df3.iloc[0:0]

"-------------FILTRO DE CUENTAS-----------"
dir_acc = r"C:\Users\rodrigo\Swiss Medical S.A\Finanzas UC - General\03. Posicion Financiera Diaria\02. Transferencias\Acl_Acc.xlsx"
xl = xw.Book(dir_acc)

contador_cuentas=17

for i in range(contador_cuentas):
    df3=df3.append(df2.loc[df2['Cuenta (Formato Interbanking)'] == xl.sheets[1].range(i+2, 2).value])

xl.close()

df3["CC Transfer"]=0

"--------NUEVA COLUMNA CON CC DEBITOS----------"

dfacc=pandas.read_excel(dir_acc, sheet_name="Hoja1")

i=0
for i in range(len(cias)):
    df3.at[df3.index[df3.loc[:,'Soc.']==cias[i][0]],"CC Transfer"] = str(dfacc["CC Debito"].values[dfacc.index[dfacc.loc[:,'Soc']==cias[i][0]]])



"---------EMPIEZA LA INTERFAZ----------------"
#RAIZ
raiz=Tk()
raiz.title("SMG GROUP // " + arrow.now().format('DD-MM-YYYY'))
raiz.geometry("1350x550")

#FONDO
img = ImageTk.PhotoImage(Image.open("SwissMedical.png"))
panel = Label(raiz, image = img)
panel.place(x=0, y=0, relwidth=1, relheight=1)

#BOTONOS Y LABELS
y=20
#for i in range(len(cias)):
#    text = Label(panel, text=cias[i][1])
#    text.place(x=30, y=y)
#    y+=30

text = tkinter.Text(width=2000)
text.insert(tkinter.END, str(df3))
text.pack(fill=X)



combo= ttk.Combobox()
combo.place(x=30, y=450)
combo["state"]="readonly"
combo["values"]=["SWISS MEDICAL",
                 "VIDA",
                 "ART",
                 "SEGUROS",
                 "ECCO",
                 "SERVICIOS",
                 "RETIRO",
                 "RE",
                 "SAN LUCAS",
                 "INVESTMENT",
                 "SERVICES",
                 "CORPORATE",
                 "GLOBAL INVESTMENT",
                 "LIBOFE",
                 "INTERNACIONAL",
                 "TRAVEL",
                 "MANUTO",
                 "INVERSIONES",
                 "MELL",
                 "MICROCENTRO.",
                 "INSTITUTO SALTA",
                 "WOLLAK",
                 "BLUE CROSS AND BLUE SHIELD",
                 "BEFERRA",
                 "FUNDACION",
                 "INMOBILIARIA",
                 "JET MATCH",
                 "MEDIA",
                 "ECCO",
                 "CIRO VENTURES",
                 "CLINICAS MENDOZA",
                 "FIDUCIARIA SAU"
]

def filtrador():
    global df3
    global text
    text.destroy()
    df4 = df3.loc[df2['Razón Social'] == combo.get()]
    text = tkinter.Text(width=2000)
    text.insert(END, str(df4))
    text.pack(fill=X)

def clear():
    global df3
    global text
    text.destroy()
    text = tkinter.Text(width=2000)
    text.insert(END, str(df3))
    text.pack(fill=X)

def printer():
    global matriz_acl
    print(matriz_acl)

#"SWISS MEDICAL S.A. ","SMG LIFE SEGUROS DE VIDA S.A.","SWISS MEDICAL ART. S.A.","SMG COMPAÑÍA ARGENTINA DE SEGUROS S.A.","ECCO S.A ","SMG SERVICIOS S.A.            ","SMG LIFE CIA. DE SEGUROS DE RETIRO S.A.","SMG RE","CONSULTORIOS INTEGRALES SAN LUCAS S.A.","SMG INVESTMENT S.A.          ","SMG SERVICES S.A.            ","SMG CORPORATE S.A.           ","GLOBAL INVESTMENT S.A.         ","LIBOFE S.A.                  ","INTERNACIONAL CIA SEG VIDA S.A.,SMG TRAVEL S.A.                ","MANUTO SEGURIDAD S.A.          ","Swiss INVERSIONES SA","MELL S.A.","Microcentro l S.A.","Instituto de Salta Cía. De Seguros de Vida S.A.","Wollak","Blue Cross & Blueshield","SWISS MEDICAL ART. S.A. (Exenta)","BEFERRA S.A.","FUNDACION SWISS MEDICAL","SMG Inmobiliaria SA","JET MACH","SMG MEDIA SA (Ex Grupo de Nervaez)","ECCO SA","CIRO VENTURES","SMG Clínicas Mendoza","SMG FIDUCIARIA SAU"

botonarchivotransfer=Button(panel, text="Aceptar", height=1, width= 24,command=lambda:filtrador())
botonarchivotransfer.place(x=30, y=500)

botonarchivotransfer2=Button(panel, text="Clear", height=1, width= 24,command=lambda:clear())
botonarchivotransfer2.place(x=250, y=500)

botonarchivotransfer2=Button(panel, text="Imprimir", height=1, width= 24,command=lambda:printer())
botonarchivotransfer2.place(x=450, y=500)

raiz.mainloop()







