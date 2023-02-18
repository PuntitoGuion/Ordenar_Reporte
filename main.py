import pandas as pd
import os
import pygame
import requests
from tkinter import filedialog, messagebox
from io import BytesIO


def sonido(url):
    """URL ejemplo: \n
    https://drive.google.com/uc?export=download&id=FILE_ID"""
    pygame.mixer.init()
    response = requests.get(url)

    audio_data = BytesIO(response.content)

    with open("temp_audio.mp3", "wb") as f:
        f.write(audio_data.getbuffer())

    pygame.mixer.music.load("temp_audio.mp3")
    pygame.mixer.music.play()

    # Esperar hasta que el audio termine de reproducirse
    while pygame.mixer.music.get_busy():
        continue
        
    pygame.quit()
    os.remove("temp_audio.mp3")

sonido("https://drive.google.com/uc?export=download&id=1VJD44d5SF3g2H5rH0CJIW1bo5DfMkDeV")


def sort_group(group):
    sorted_group = group.sort_values(by=["Fecha", "Cuenta"], ascending=[True,False])
    return sorted_group



dirFile = filedialog.askopenfilename(title="Reporte faltante", filetypes=(("Archivos Excel", "*.xls"), ("Archivos Excel", "*.xlsx"), ("Todos los archivos", "*.*")))

fileName = dirFile.split("/")[-1].split(".")[0]

reportes = pd.read_csv(dirFile, delimiter="\t", encoding="ISO-8859-1")

cant_reportes = reportes.shape[0]

archivosFaltantes = reportes["Fecha"][cant_reportes - 10]

totalLocalesFaltantes = reportes["Fecha"][cant_reportes - 7]


# Manipulacion de datos
reportes = pd.read_csv(dirFile, delimiter="\t", encoding="ISO-8859-1", nrows=cant_reportes - 10)
reportes['Fecha'] = pd.to_datetime(reportes["Fecha"], format='%d/%m/%Y')
reportes['Fecha'] = reportes['Fecha'].dt.strftime('%d/%m/%Y')
reportes['Local'] = reportes['Local'].astype(str)
reportes = reportes.sort_values(['Local', 'Fecha'])
first_dates = reportes.groupby('Local')['Fecha'].min().reset_index()
first_dates = first_dates.rename(columns={'Fecha': 'first_date'})
df = pd.merge(reportes, first_dates, on='Local')
df['Cuenta'] = df.groupby('Local')['Local'].transform('count')
df = df.sort_values(['first_date', 'Local', 'Fecha', 'Cuenta']).reset_index(drop=True)
df = df.groupby('Local').apply(sort_group).reset_index(drop=True)
df = df.drop('first_date', axis=1)
nuevo_registro = pd.DataFrame({'Local': ['Archivos Faltantes: ','Total de Locales c/Faltantes :'], 'Estado': [archivosFaltantes,totalLocalesFaltantes],'Cuenta':['',''],'Tipo de Informacion':['',''],'Fecha':['',''],})
df = df.append(nuevo_registro,ignore_index=True)

# Guardar y abrir Excel

while True:
    try:
        df.to_excel(f"{fileName}.xlsx",index=False)
        sonido("https://drive.google.com/uc?export=download&id=1XunfCmrO_6BHtYcoxL6Cyq4HuE66UlEs")
        ruta_actual = os.getcwd()
        os.chdir(ruta_actual)
        os.system(f'start excel.exe "{fileName}.xlsx"')
        break
    except PermissionError:
        messagebox.showerror("¡Error!","Favor de mantener cerrado la hoja de Excel para poder guardar la información")