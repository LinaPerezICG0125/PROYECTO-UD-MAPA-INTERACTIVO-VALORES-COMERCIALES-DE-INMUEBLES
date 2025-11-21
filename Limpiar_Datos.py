import os
import re
import pandas as pd

archivo_entrada = "data/Datos Proyecto.xlsx"
archivo_salida = "data/Dirección Limpia.xlsx"
col_origen = "DIRECCIÓN"
col_salida = "Dirección Estandarizada"

df = pd.read_excel(archivo_entrada)
print("Archivo cargado correctamente.")
print("Primeras filas:")
print(df.head())

def nom_lim(nom):
    if not isinstance(nom, str):
        return ""
    nom = nom.strip()
    nom = " ".join(nom.split())
    return nom

    nom = nom.replace("No.", "#").replace("No.", "#")
    nom = nom.replace("No", "#").replace("No", "#")
    nom = nom.replace("N°", "#").replace("N°", "#")
    nom = re.sub(r'[.,;:]', ' ', nom) #Elimina caracteres innecesarios
    nom = re.sub(r'\s*-\s*', '-', nom) #Elimina espacios antes y despúes de -
    nom = re.sub(r'\s+', ' ', nom) #Elimina espacios entre caracteres

    nom = re.sub(r'\b(kr|Kr|Kr\.|kr\.)\b', 'carrera', nom)
    nom = re.sub(r'\b(cr|Cr|Cra|cra|cr\.|Cr\.|Cra\.|cra\.)\b', 'carrera', nom)
    nom = re.sub(r'\b(cll|Cll|cl|Cl|cll\.|Cll\.|cl\.|Cl\.)\b', 'calle', nom)
    nom = re.sub(r'\b(av|Av|AV|av\.|Av\.|AV\.)\b', 'avenida', nom)
    nom = re.sub(r'\b(dg|Dg|DG|dg\.|Dg\.|DG\.)\b', 'diagonal', nom)
    nom = re.sub(r'\b(tv|Tv|TV|Transv|tv\.|Tv\.|TV\.|Transv\.)\b', 'transversal', nom)
    nom = re.sub(r'\b(ac|Ac|AC|ac\.|Ac\.|AC\.)\b', 'avenida calle', nom)
    nom = re.sub(r'\b(ak|Ak|AK|ak\.|Ak\.|AK\.)\b', 'avenida carrera', nom)
    nom = re.sub(r'\s+', ' ', nom).strip()

    if '#' in nom:
        nom = re.sub(r'\s*#\s*', '#', nom) #Elimina espacios antes y después de #
        nom = re.sub(r'\s+', ' ', nom).strip()
    else:
        nom = re.sub(r'(\d+[a-z]?)\s+(\d)', r'\1 # \2', nom)
    nom = nom.title()

if col_origen not in df.columns:
    print("Columnas disponibles:", df.columns.tolist())
    raise ValueError (f"No se encontró columna de dirección '{col_origen}'")

df["DIRECCIÓN"] = df["DIRECCIÓN"].astype(str)     
df[col_origen] = df[col_origen].astype(str)
df[col_origen] = df[col_origen].replace('nan', '')
df[col_origen] = df[col_origen].str.strip()
df[col_origen] = df[col_origen].str.replace(r'\s+', ' ', regex=True)

df[col_salida] = df[col_origen].apply(nom_lim)
print("Muestra original:")
print(df[[col_origen, col_salida]].head(15))
os.makedirs(os.path.dirname(archivo_salida) or ",", exist_ok=True)
df.to_excel(archivo_salida, index=False, engine="openpyxl")
print("Archivo guardado en:", archivo_salida)