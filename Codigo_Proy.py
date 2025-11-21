import re
from geopy.geocoders import Nominatim
import pandas as pd
import time
import folium

archivo = "Datos Proyecto.xlsx"
df = pd.read_excel(archivo)
print("Archivo cargado correctamente.")
print("Primeras filas:")
print(df.head())

def nom_lim(direccion):
    if not isinstance(direccion, str):
        return " "
    nom = direccion.lower().strip()

    nom = nom.replace("No.", "#").replace("No.", "#")
    nom = nom.replace("No", "#").replace("No", "#")
    nom = re.sub(r'[.,:]', ' ', nom)
    nom = re.sub(r'\s*-\s*', '-', nom)

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
        nom = re.sub(r'\s*#\s*', '#', nom)
        nom = re.sub(r'\s+', ' ', nom).strip()
    else:
        nom = re.sub(r'(\d+[a-z]?)\s+(\d)', r'\1 # \2', nom)
    nom = nom.title()
    if "Bogotá" not in nom and "Bogota" not in nom:
        nom += ", Bogotá, Colombia"
    return nom

def coordenadas(direccion, geolocator):
    try:
        nom_full = nom_lim(direccion)
        location = geolocator.geocode(nom_full, timeout=10)
        if location:
            print(f"{nom_full}")
            print(f"Latitud: {location.latitude}, Longitud:{location.longitude}")
            print("-" * 60)
            return location.latitude, location.longitude
        else:
            print(f"No se encontró: {nom_full}")
            print("-" * 60)
            return None
    except Exception as e:
        print(f"Error con:'{direccion}':{e}")
        print("-" * 60)
        return None
    
geolocator = Nominatim(user_agent="Datos Proyecto.xlsx")
resultados = []
no_encontrados = []

for d in df["DIRECCIÓN"]: #Nombre columna
    d_full = nom_lim(d)
    resultado = coordenadas(d_full, geolocator)
    if resultado:
        resultados.append({"Dirección_original": d, "Dirección_limpia": d_full, "Latitud": resultado[0], "Longitud": resultado[1]})
    else:
        no_encontrados.append(d)
    time.sleep(1)

if no_encontrados:
    with open("Direcciones no encontradas.txt", "w", encoding="utf-8") as f:
        for d in no_encontrados:
            f.write(str(d) + "\n")
        print(f"\n Se guardaron{len(no_encontrados)} direcciones no encontradas en 'Direcciones no encontradas.txt'")
else:
    print("\n Todas las direeciones fueron encontradas correctamente.")

if resultados:
    df_resultados = pd.Dataframe(resultados)
    df_resultados.to_csv("Coordenadas_direcciones.csv", index=False, encoding="utf-8-sig")
    df_resultados.to_excel("Coordenadas_direcciones.xlsx", index=False)
    print("Datos guardados en 'Coordenadas_direcciones.csv' y 'Coordenadas_direcciones.xlsx'")
else:
    print("No se encontraron coordenadas para guardar.")