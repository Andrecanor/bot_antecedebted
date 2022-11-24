from os import mkdir
from os.path import join, exists
import datetime as dt
import pandas as pd


work_path = GetVar("work_path")
#db_name = r"C:\Users\aprendiz8.automatiza\OneDrive - GANA S.A\Seleccion MVP - Plan Carrera\PlanCarrera.xlsx"

today = dt.datetime.today()

folder_pdfs  = join(work_path, "Antecedentes")
if not exists(folder_pdfs):
    mkdir(folder_pdfs)

folder_mounth = join(folder_pdfs, today.strftime("%Y-%B"))
if not exists(folder_mounth):
    mkdir(folder_mounth)

folder_today = join(folder_mounth, f"pdfs del {today.strftime('%d')}" )
if not exists(folder_today):
    mkdir(folder_today)


with open(db_name, mode='rb') as file:
    df = pd.read_excel(file, sheet_name=0, engine='openpyxl', dtype={"Documento ":str, "Fecha de postulacion":str})

new_df = df[df["Fecha de postulacion"].str.contains(today.strftime("%Y-%m-%d"))]


list_ID = new_df["Documento "].to_list()


SetVar("users_ID", list_ID)
SetVar("folder_today", folder_today)
