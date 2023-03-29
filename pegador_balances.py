# noinspection PyUnresolvedReferences
import arrow

# noinspection PyUnresolvedReferences
import os

# noinspection PyUnresolvedReferences
import xlwings as xw

# noinspection PyUnresolvedReferences
import pandas

# noinspection PyUnresolvedReferences
import ctypes

pandas.options.display.width = 0

dir=r"C:\Users\rodriaguirre\Desktop\blcs07.xlsx"
dir_info=r"C:\Users\rodriaguirre\OneDrive - Swiss Medical S.A\Documents\General\02. Reportes Corporativos\01. Informe Financiero\2022\3700 MELL\INFFIN Mensual 09 2022 MELL.xlsx"

df = pandas.read_excel(dir)
df= df.iloc[:,2:16]
df.rename(columns={"Unnamed: 2": "Soc","Unnamed: 4": "Mayor","Unnamed: 14": "Variacion", "Unnamed: 7": "Descripcion", "Unnamed: 10": "T+0", "Unnamed: 12": "T-1"}, inplace=True)
df=df.drop(["Unnamed: 15"], axis=1)
df.dropna(inplace=True, subset=['Soc'])
df["data_type"]= df["Soc"].apply(lambda x: "STR" if isinstance(x, str) else "num")
df=df.loc[df["data_type"]=="num"]





writer.save()

print(df)

