# noinspection PyUnresolvedReferences
from tkinter import *

# noinspection PyUnresolvedReferences
import arrow

# noinspection PyUnresolvedReferences
from PIL import ImageTk, Image

# noinspection PyUnresolvedReferences
import os

# noinspection PyUnresolvedReferences
from armados import armadora

# noinspection PyUnresolvedReferences
from main import transferencias

# noinspection PyUnresolvedReferences
import xlwings

# noinspection PyUnresolvedReferences
import pandas

# noinspection PyUnresolvedReferences
import ctypes

# noinspection PyUnresolvedReferences
import locale

# noinspection PyUnresolvedReferences
import mails

# noinspection PyUnresolvedReferences
import armados

# noinspection PyUnresolvedReferences
from auxiliar import auxiliar3

# noinspection PyUnresolvedReferences
import loggins

locale.setlocale(locale.LC_ALL, 'en_US')

#RAIZ
raiz=Tk()
raiz.title("SMG GROUP // " + arrow.now().format('DD-MM-YYYY'))
raiz.geometry("650x350")

#FONDO
img = ImageTk.PhotoImage(Image.open("SwissMedical.png"))
panel = Label(raiz, image = img)
panel.place(x=0, y=0, relwidth=1, relheight=1)

#LABELS
text2= Label(panel, text= "Estado Bajada", font=("Arial", 7))
text2.grid(row=0, column=3)

text4= Label (panel, text=" ")
text4.grid(row=0, column=1, padx=145)

text5 = Label(panel, text="Débito Directo")
text5.grid(row=2, column=2, padx=15, columnspan=3)

text6 = Label(panel, text="Generales: ")
text6.grid(row=3, column=2, columnspan=1)

text7 = Label(panel, text="----------", relief="sunken", width=12)
text7.grid(row=3, column=3, columnspan=2)

text8 = Label(panel, text="Life: ")
text8.grid(row=4, column=2, columnspan=1)

text9 = Label(panel, text="----------", relief="sunken", width=12)
text9.grid(row=4, column=3, columnspan=2)

#BOTONES
botonarchivotransfer=Button(panel, text="Crear Archivo de Transferenicas", height=1, width= 24,command=armadora.armadorArchTransfer)
botonarchivotransfer.grid(row=1,column=0, padx=5, pady=9)

botonarchivotransfer=Button(panel, text=">", height=1, width= 6,command=lambda: mails.mailTransfer())
botonarchivotransfer.place(x=190,y=29)

botonarchivotransfer=Button(panel, text=">", height=1, width= 6,command=lambda: mails.mailOps())
botonarchivotransfer.place(x=411,y=118)

botonarchivotransfer=Button(panel, text=">", height=1, width= 6,command=lambda: mails.mailAcl())
botonarchivotransfer.place(x=411,y=73)

botonarchivoextracto=Button(panel, text="Filtrar Extractos DATANET",height=1, width= 24,command=lambda: armadora.armadorArchExtractos())
botonarchivoextracto.grid(row=4,column=0, padx=1, pady=9)

botonarchivoextracto=Button(panel, text="Abrir Consolidado",height=1, width= 24,command=armadora.aperturaConsolidado)
botonarchivoextracto.grid(row=3,column=0, padx=1, pady=9)

botonarchivoextracto=Button(panel, text="Operaciones",height=1, width= 18, command=lambda: armadora.operaciones())
botonarchivoextracto.grid(row=3,column=1, padx=1, pady=9)

botonarchivoextracto=Button(panel, text="Armar Transferencias",height=1, width= 24,command=lambda: transferencias.armadorTransferencias())
botonarchivoextracto.grid(row=2,column=0, padx=1, pady=9)

botonarchivoextracto=Button(panel, text="Armar Saldia",height=1, width= 24,command=lambda: armadora.armadorPosicion())
botonarchivoextracto.grid(row=7,column=0, padx=1, pady=9)

botonarchivopagos=Button(panel, text="Armar Pagos", height=1, width= 24,command=lambda: armadora.armadorPagos())
botonarchivopagos.grid(row=6,column=0, padx=1, pady=9)

botonarchivopagos=Button(panel, text="Armar Extractos para Cash", height=1, width= 24,command=lambda: armadora.armadorExtractoscash())
botonarchivopagos.grid(row=5,column=0, padx=1, pady=9)

botonarchivopagos=Button(panel, text="Refresh", height=1, width= 18,command=lambda: directo())
botonarchivopagos.grid(row=7,column=2, padx=1, pady=9, columnspan=3)

botonarchivopagos=Button(panel, text="Buscar DD", height=1, width= 18,command=lambda: ddBuscador())
botonarchivopagos.grid(row=5,column=2, padx=1, pady=9, columnspan=3)

botonarchivopagos=Button(panel, text="Cash para UN", height=1, width= 18,command=lambda: armadora.transferirFlujos())
botonarchivopagos.grid(row=1,column=1, padx=1, pady=9)

botonarchivopagos=Button(panel, text="Aclaraciones", height=1, width= 18,command=lambda: transferencias.transfer_Acl())
botonarchivopagos.grid(row=2,column=1, padx=1, pady=9)

botonarchivopagos=Button(panel, text="Propuestas Pagos", height=1, width= 18,command=lambda: armadora.propuestas_paula())
botonarchivopagos.grid(row=4,column=1, padx=1, pady=9)

botonarchivoextracto=Button(panel, text="Login IB",height=1, width= 18,command=lambda: loggins.log_IB())
botonarchivoextracto.grid(row=7,column=1, padx=1, pady=9)

botonarchivoextracto=Button(panel, text="Mails Contrapartes",height=1, width= 18,command=lambda: mails.mailContrapartes())
botonarchivoextracto.grid(row=5,column=1, padx=1, pady=9)

botonarchivoextracto=Button(panel, text="Out",height=1, width= 18,command=lambda: loggins.out())
botonarchivoextracto.grid(row=6,column=1, padx=1, pady=9)

#IMAGENES PARA CONTROL MARGEN DERECHO
img2 = ImageTk.PhotoImage(Image.open("neutro.png").resize((14,14), Image.ANTIALIAS))
panel2 = Label(panel, image = img2)
panel2.grid(row=1, column=3, padx=2.5)

def directo():
    global panel2
    imgok = ImageTk.PhotoImage(Image.open("ok.png").resize((14, 14), Image.ANTIALIAS))
    imgerror = ImageTk.PhotoImage(Image.open("error.png").resize((14, 14), Image.ANTIALIAS))
    if (os.path.isfile(r"C:\Users\rodriaguirre\OneDrive - Swiss Medical S.A\Documents\General\03. Posicion Financiera Diaria\06. Pagos SAP\zcl04 " + arrow.now().format('DD.MM') + ".XLS")):
        panel2.configure(image=imgok)
        panel2.photo=imgok
    else:
        panel2.configure(image=imgerror)
        panel2.photo=imgerror

def ddBuscador():
    global text7
    dir = r"C:\Users\rodriaguirre\Desktop\extractos_mult.xlsx"

    #GENERALES
    xlext = pandas.read_excel(dir, sheet_name="Generales", header=None)
    garray = xlext.index[xlext.iloc[:, 15] == "DEBITO DIRECTO - RENDICIO"]

    if (len(garray) != 0):
        listg = []
        for f in range(len(garray)):
            listg.append([])
            for c in range(1):
                listg[f].append(garray[f] + 1)
                listg[f].append("-")
                listg[f].append(locale.format_string("%d", xlext.iloc[garray[f], 8], grouping=True))
        text7.configure(text=listg, width=12, bg="#00FF00", relief="sunken")
        ctypes.windll.user32.MessageBoxW(0, "EMERGENCIA ENTRA DEBITO DIRECTO!!! ENTRA DEBITO DIRECTO ATENCION!!! ATENCION!!!", "Confirmación", 0)
    else:
        text7.configure(text="----------", bg="#ff0000", width=12, relief="sunken")

    #LIFE
    xlext = pandas.read_excel(dir, sheet_name="Life", header=None)
    garray = xlext.index[xlext.iloc[:, 15] == "RECAUDACION DEB.DIR. A C"]

    if (len(garray) != 0):
        listg = []
        for f in range(len(garray)):
            listg.append([])
            for c in range(1):
                listg[f].append(garray[f] + 1)
                listg[f].append("-")
                listg[f].append(locale.format_string("%d", xlext.iloc[garray[f], 8], grouping=True))
        text9.configure(text=listg, width=12, bg="#00FF00", relief="sunken")
        ctypes.windll.user32.MessageBoxW(0, "EMERGENCIA ENTRA DEBITO DIRCTO!!! ENTRA DEBITO DIRECTO ATENCION!!! ATENCION!!!", "Confirmación", 0)
    else:
        text9.configure(text="----------", bg="#ff0000", width=12, relief="sunken")

raiz.mainloop()


