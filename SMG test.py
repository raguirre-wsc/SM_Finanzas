# noinspection PyUnresolvedReferences
import tkinter

import pandas
# noinspection PyUnresolvedReferences
import arrow
# noinspection PyUnresolvedReferences
from auxiliar import auxiliar
# noinspection PyUnresolvedReferences
from auxiliar import auxiliar2
# noinspection PyUnresolvedReferences
from auxiliar import auxiliar3
# noinspection PyUnresolvedReferences
import ctypes
# noinspection PyUnresolvedReferences
import xlwings
# noinspection PyUnresolvedReferences
import tkinter as tkinter
# noinspection PyUnresolvedReferences
import webbrowser

pandas.set_option('display.colheader_justify', 'center')

dir_t = "C:/Users/rodriaguirre/OneDrive - Swiss Medical S.A/Documents/General/03. Posicion Financiera Diaria/02. Transferencias/" + arrow.now().format('YYYY') + "/" + arrow.now().format('MM') + ". " + auxiliar.nombreMes(str(arrow.now().format('MM'))) + "/Transferencias " + arrow.now().format('DD-MM-YYYY') + ".xlsm"
dir_sg= "C:/Users/rodriaguirre/OneDrive - Swiss Medical S.A/Documents/General/03. Posicion Financiera Diaria/01. Saldia/02. Seguros/" + arrow.now().format('YYYY') + "/" +arrow.now().format('MM.YYYY') + "/" + arrow.now().format('YYYY_MM_DD') + " SMG.xlsx"
dir_ss= r"C:\Users\rodriaguirre\OneDrive - Swiss Medical S.A\Documents\General\03. Posicion Financiera Diaria\01. Saldia\01. Salud\Saldia.xlsx"

#SALDIA SEGUROS
dfs = pandas.read_excel(dir_sg, sheet_name="SALDIA")
montoseg=dfs.iloc[0,19]
dfs = dfs.iloc[:,[23,21]]
dfs=dfs.loc[dfs["imps"]!=0]
dfs.dropna(inplace=True, subset=['imps'])
dfs = dfs.astype({"imps":"int64","cod":"int"})
#print(dfs)

#SALDIA SALUD
dfsa = pandas.read_excel(dir_ss, sheet_name="Posicion Financiera")
montosal=dfsa.iloc[53,12]*1000+dfsa.iloc[106,12]*1000
dfsa = dfsa.iloc[:,[36,12]]
dfsa=dfsa.loc[dfsa["imps"]!=0]
dfsa.dropna(inplace=True, subset=['imps'])
dfsa.dropna(inplace=True, subset=['cod'])
dfsa = dfsa.astype({"imps":"int64","cod":"int"})
#dfsa["imps"] = abs(dfsa["imps"])
dfsa["imps"] = dfsa["imps"]*1000
dfsa["imps"]=dfsa["imps"]*-1
#print(dfsa)

#APPEND
dfa=pandas.concat([dfs, dfsa])
#print(dfa)

#TRANSFERENCIAS
#BASE DE CUENTAS
dftb = pandas.read_excel(dir_t, sheet_name="Base")
dftb = dftb.iloc[:,[15,16]]
dftb.rename(columns = {'Concatenado.1':'cuenta'}, inplace = True)

#OPERACIONES
dft = pandas.read_excel(dir_t, sheet_name="Transferencias")
montot=dft.iloc[841,3]
dft = dft.iloc[:,[14,10]]
dft=dft.loc[dft["impt"]!=0]
dft.dropna(inplace=True, subset=['impt'])
dft.dropna(inplace=True, subset=['cod'])
dft["impt"]=dft["impt"]*-1
dft = dft.astype({"impt":"int64","cod":"int"})
dft= dft.groupby(['cod']).sum()


#MATRIZ CONTROL DFM
def control(x):
    return "OK" if x["impt"] == x["imps"] else "ERROR"
dfm=dft.merge(dfa,on="cod", how="outer")
dfm=dfm.merge(dftb,on="cod", how="left")
dfm["control"]= dfm.apply(control, axis=1)
#print(f"{type(dfm.shape[0])} and {type(dfm.shape[1])}")

#MONTOS TRANSFERIDOS
montos=montosal+montoseg
montot=montot
print(dfm)

#dmf=dfm.T.reset_index().T

x = dfm.shape[0]
y = dfm.shape[1]
libro = open("C:/Users/rodriaguirre/Desktop/contol_transfer.html", "w")
html = f"""
       <h1>Transferencias {arrow.now().format('DD.MM.YYYY')}</h1>
       <table>
       <tr>
       <th style="text-align:left; background-color:#3a6070; color:#FFF">Cod</th>
       <th style="text-align:left; background-color:#3a6070; color:#FFF">Monto Transferencia</th>
       <th style="text-align:left; background-color:#3a6070; color:#FFF">Monto Saldia</th>
       <th style="text-align:left; background-color:#3a6070; color:#FFF">Cuenta</th>
       <th style="text-align:left; background-color:#3a6070; color:#FFF">Control</th>
       </tr>
       """

if dfm.loc[dfm["control"]=="error"].shape[0]==0 and montos==montot: #si el conteo de la palabra "error" = 0 en la col "control y los montos totales son iguales
    for j in range(x):
        html_append = f"""
               <tr>
               <td style="border:1px solid #e3e3e3; padding:4px 8px">{str(dfm.at[j,"cod"])}</td>
               <td style="border:1px solid #e3e3e3; padding:4px 8px; font-weight: bold">{'{:,.0f}'.format(dfm.at[j,"impt"])}</td>
               <td style="border:1px solid #e3e3e3; padding:4px 8px; font-weight: bold">{'{:,.0f}'.format(dfm.at[j,"imps"])}</td>
               <td style="border:1px solid #e3e3e3; padding:4px 8px">{dfm.at[j,"cuenta"]}</td>
               <td style="border:1px solid #e3e3e3; padding:4px 8px; color:green; font-weight: bold">{dfm.at[j,"control"]}</td>
               </tr>
               """
        html+=html_append
    html_append = f"""
    </table> """
    html += html_append
    ctypes.windll.user32.MessageBoxW(0, "OK - OK - OK - OK", "CONFIRMACION", 0)
else:
    for j in range(x):
        html_append = f"""
               <tr>
               <td style="border:1px solid #e3e3e3; padding:4px 8px">{str(dfm.at[j,"cod"])}</td>
               <td style="border:1px solid #e3e3e3; padding:4px 8px; font-weight: bold">{str('{:,.0f}'.format(dfm.at[j,"impt"]))}</td>
               <td style="border:1px solid #e3e3e3; padding:4px 8px; font-weight: bold">{str('{:,.0f}'.format(dfm.at[j,"imps"]))}</td>
               <td style="border:1px solid #e3e3e3; padding:4px 8px">{str(dfm.at[j,"cuenta"])}</td>"""
        html += html_append
        if dfm.at[j, "control"]=="OK":
            html_append = f"""
               <td style="border:1px solid #e3e3e3; padding:4px 8px; color:green; font-weight: bold">{str(dfm.at[j,"control"])}</td>
               </tr>"""
            html+=html_append
        else:
            html_append = f"""
                <td style="border:1px solid #e3e3e3; padding:4px 8px; color:red; font-weight: bold">{str(dfm.at[j,"control"])}</td>
                </tr>"""
            html += html_append
    html_append = f"""
    </table> """
    html += html_append
    ctypes.windll.user32.MessageBoxW(0, "Error - Error - Error - Error", "Alerta", 0)
libro.write(html)
