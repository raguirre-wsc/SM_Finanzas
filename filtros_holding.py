# noinspection PyUnresolvedReferences
from openpyxl import Workbook

# noinspection PyUnresolvedReferences
import ctypes

# noinspection PyUnresolvedReferences
from tkinter import *

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
from openpyxl import load_workbook

# noinspection PyUnresolvedReferences
import xlsxwriter

# noinspection PyUnresolvedReferences
import keyboard

# noinspection PyUnresolvedReferences
import time

# noinspection PyUnresolvedReferences
import os


def filtro():
    matriz = [
        [3900, "FUNDACION"],
        [3300, "MANUTO"],
        [3500, "LIBOFE"],
        [3600, "JET MATCH"],
        [3700, "MELL"],
        [3800, "MICROCENTRO"],
        [3400, "SAINT GALL"],
        [8200, "SMG CORPORATE"],
        [8000, "SMG INVESTMENT"],
        [6500, "SMG SERVICES"],
        [3000, "SMG SERVICIOS"],
        [8100, "SWISS INVERSIONES"],
        [7600, "INTERNACIONAL"]
    ]

#    dir = "C:/Users/rodrigo/Swiss Medical S.A/Finanzas UC - General/02. Reportes Corporativos/01. Informe Financiero/FLQLS Holdings/2021/09/flqls.xlsx"

    dir_blcs= r"C:\Users\rodriaguirre\Desktop\blcs07.xlsx"
    blcs=pandas.read_excel (dir_blcs)

    blcs = blcs.iloc[:, 2:16]
    blcs.rename(columns={"Unnamed: 2": "Soc", "Unnamed: 4": "Mayor", "Unnamed: 14": "Variacion", "Unnamed: 7": "Descripcion", "Unnamed: 10": "T+0", "Unnamed: 12": "T-1"}, inplace=True)
    blcs = blcs.drop(["Unnamed: 15"], axis=1)
    blcs.dropna(inplace=True, subset=['Soc'])
    blcs["data_type"] = blcs["Soc"].apply(lambda x: "STR" if isinstance(x, str) else "num")
    blcs = blcs.loc[blcs["data_type"] == "num"]

    writer = pandas.ExcelWriter(r"C:\Users\rodriaguirre\OneDrive - Swiss Medical S.A\Documents\General\02. Reportes Corporativos\01. Informe Financiero\FLQLS Holdings\2023\09\abc.xlsx", engine='xlsxwriter')

    for i in range(13):
        # FILTRADOR DE EXTRAC
        blcs_filtrada = blcs.loc[blcs["Soc"] == matriz[i][0]]

        # Write each dataframe to a different worksheet.
        blcs_filtrada.to_excel(writer, sheet_name=matriz[i][1], index=False, header=False)

# Close the Pandas Excel writer and output the Excel file.
    writer.save()

    ctypes.windll.user32.MessageBoxW(0, "Proceso termiando", "Confirmación", 0)

filtro()

print("corri")
