import pandas as pd

repor = pd.read_excel("Reporte.xlsx")
diario = pd.read_excel("ReporteDiario25-10-2023.xlsx")
archivoG = pd.read_excel("Base_Completa.xlsx")


df_1 = pd.merge(repor,archivoG,how="left", left_on="Enlace", right_on="Nombre_Orion")
df_1 = pd.concat([diario,df_1])
df_1= df_1.drop_duplicates(keep="first",subset=["Enlace","Tiempos"])
df_1.info()
df_1.to_excel("ReporteActualizado.xlsx")
df_1.info()



