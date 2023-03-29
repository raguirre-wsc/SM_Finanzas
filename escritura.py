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
import openpyxl

# noinspection PyUnresolvedReferences
from auxiliar import auxiliar

# noinspection PyUnresolvedReferences
from auxiliar import auxiliar2

# noinspection PyUnresolvedReferences
from auxiliar import auxiliar3

# noinspection PyUnresolvedReferences
import time

def TransAM():
    libro = open(r"C:\Users\rodriaguirre\Desktop\libro.html","w")


    """Transferencias"""
    date="C:/Users/rodriaguirre/OneDrive - Swiss Medical S.A/Documents/General/03. Posicion Financiera Diaria/02. Transferencias/" + arrow.now().format('YYYY') + "/" + arrow.now().format('MM') + ". " + auxiliar.nombreMes(str(arrow.now().format('MM'))) + "/Transferencias " + arrow.now().format('DD-MM-YYYY') + ".xlsm"
    xl= xlwings.Book(date)
    xls=xl.sheets['Transferencias']

    text=""

    array=[
        ["A451","D451","A452","E452"],  #art - 1
        ["A454","D454","A455","E455"],  #art - 2
        ["A81","D81","A82","E82"],  #seguros
        ["A7","D7","A8","E8"],  #swiss
        ["A44","D44","A45","E45"],   #ecco
        ["A488","D488","A489","E489"],   #life
        ["A562","D562","A563","E563"]   #retiro
    ]

    virgen=0

    for i in array:
        try:
            bco_bna=xls.range(i[0]).value
            bco_icbc1=xls.range(i[2]).value

            imp_bna=int(xls.range(i[1]).value)
            imp_icbc1=int(xls.range(i[3]).value)

            if virgen==0:
                BNA_ICBC= f"\n<b>{bco_bna}&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;{imp_bna:,} <br>\n{bco_icbc1}&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;{imp_icbc1:,}</b>"
                virgen=1
                text += BNA_ICBC
            else:
                BNA_ICBC = f"<br>\n<br>\n<b>{bco_bna}&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;{imp_bna:,} <br>\n{bco_icbc1}&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;{imp_icbc1:,}</b>"
                text += BNA_ICBC
        except:
            text += ""

    html=f"""
    Estimados,<br>
    <br>
    Agregamos transferencias urgentes: <br>
    <br>
    <p>{text}
    </p>
    Saludos.<br>
    <br>
    <IMG SRC="data:image/png;base64,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">
    """

    libro.write(html)

    xl.close()

def mailsContra():
    dir_op=r"//xfs/rulero/Sector/SMG_Administracion_y_Finanzas/Inversiones/Estructura_Vieja/Gerencia de Inversiones/Front Office New/Liquidez/Operaciones/Monitor Operaciones 221206.xlsm"
    pd=pandas.read_excel(dir_op, sheet_name="Nueva hoja soporte")
    pd["MONTO"]=pd["MONTO"].astype('float').round(2)
    pd["OBSERVACIÓN"].fillna("-", inplace=True)

    contra=list(pd['CONTRAPARTE'].unique())



    for i in contra:
        sub_pd=pd.loc[pd['CONTRAPARTE'] == str(i)].reset_index()
        loop=pd.loc[pd['CONTRAPARTE'] == str(i)].shape[0]
        libro = open("C:/Users/rodriaguirre/Desktop/Mails/" + str(i) + ".html", "w")
        html = f"""
        Estimados, les paso las operaciones de fondos para hoy. Por favor, confirmar recepción.<br>
        <br>
        <table>

        <tr>
        <th style="text-align:left; background-color:#3a6070; color:#FFF">FECHA OP</th>
        <th style="text-align:left; background-color:#3a6070; color:#FFF">FECHA LIQ</th>
        <th style="text-align:left; background-color:#3a6070; color:#FFF">OPERACIÓN</th>
        <th style="text-align:left; background-color:#3a6070; color:#FFF">COMPAÑÍA</th>
        <th style="text-align:left; background-color:#3a6070; color:#FFF">FONDO / TÍTULO</th>
        <th style="text-align:left; background-color:#3a6070; color:#FFF">MONTO</th>
        <th style="text-align:left; background-color:#3a6070; color:#FFF">MONEDA</th>
        <th style="text-align:left; background-color:#3a6070; color:#FFF">OBSERVACIÓN</th>
        </tr>
        """

        for k in range(loop):
            if sub_pd.at[k, "OPERACIÓN"]=="RESCATE":
                html_append = f"""
                <tr>
                <td style="border:1px solid #e3e3e3; padding:4px 8px">{sub_pd.at[k, "FECHA OP"].strftime("%d/%m/%Y")}</td>
                <td style="border:1px solid #e3e3e3; padding:4px 8px">{sub_pd.at[k, "FECHA LIQ"].strftime("%d/%m/%Y")}</td>
                <td style="border:1px solid #e3e3e3; padding:4px 8px; color:red; font-weight: bold">{str(sub_pd.at[k, "OPERACIÓN"])}</td>
                <td style="border:1px solid #e3e3e3; padding:4px 8px">{str(sub_pd.at[k, "COMPAÑÍA"])}</td>
                <td style="border:1px solid #e3e3e3; padding:4px 8px">{str(sub_pd.at[k, "FONDO / TÍTULO"])}</td>
                <td style="border:1px solid #e3e3e3; padding:4px 8px">{'{:,.0f}'.format(sub_pd.at[k, "MONTO"])}</td>
                <td style="border:1px solid #e3e3e3; padding:4px 8px">{str(sub_pd.at[k, "MONEDA"])}</td>
                <td style="border:1px solid #e3e3e3; padding:4px 8px">{str(sub_pd.at[k, "OBSERVACIÓN"])}</td>
                </tr>
                """
                html += html_append
            else:
                html_append=f"""
                <tr>
                <td style="border:1px solid #e3e3e3; padding:4px 8px">{sub_pd.at[k,"FECHA OP"].strftime("%d/%m/%Y")}</td>
                <td style="border:1px solid #e3e3e3; padding:4px 8px">{sub_pd.at[k,"FECHA LIQ"].strftime("%d/%m/%Y")}</td>
                <td style="border:1px solid #e3e3e3; padding:4px 8px">{str(sub_pd.at[k,"OPERACIÓN"])}</td>
                <td style="border:1px solid #e3e3e3; padding:4px 8px">{str(sub_pd.at[k,"COMPAÑÍA"])}</td>
                <td style="border:1px solid #e3e3e3; padding:4px 8px">{str(sub_pd.at[k,"FONDO / TÍTULO"])}</td>
                <td style="border:1px solid #e3e3e3; padding:4px 8px">{'{:,.0f}'.format(sub_pd.at[k, "MONTO"])}</td>
                <td style="border:1px solid #e3e3e3; padding:4px 8px">{str(sub_pd.at[k,"MONEDA"])}</td>
                <td style="border:1px solid #e3e3e3; padding:4px 8px">{str(sub_pd.at[k,"OBSERVACIÓN"])}</td>
                </tr>
                """
                html+=html_append


        html_append="""
        </table>
        <br>
        Saludos.
        """
        html += html_append

        libro.write(html)



