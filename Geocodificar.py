import time
import pandas as pd
from geopy.geocoders import Nominatim
import folium

def geocodificador(consulta:str, geolocator:Nominatim, pausa:float=1.0):
    if not consulta or str(consulta).strip() == "":
        return None, None
    try:
        cons = consulta
        if "Bogotá" not in cons.lower() and "bogotá" not in cons.lower():
            cons = cons+", Bogotá, Colombia"
            loc = geolocator.geocode(cons, timeout=10)
            time.sleep(pausa)
        if loc:
            return loc.latitude, loc.longitude
        return None, None
    except Exception as e:
        print(f"Error en la geocodificación '{cons}': {e}")
        time.sleep(pausa)
        return None, None

def procesar_coord(doc_limpio: str, doc_coord_exc:str = "Coordenadas.xlsx", mapa_int:str = "Mapa.htlm", columna_lim:str = "Dirección Estandarizada"):
    df = pd.read_excel(doc_limpio)
    if columna_lim not in df.columns:
        raise ValueError(f"No se encontro la columna '{columna_lim}' en {doc_limpio}")
       
    geolocator = Nominatim(user_agent="Dirección Limpia.xlsx")
    df["Latitud"] = None
    df["Longitud"] = None
    no_encontradas = []

    for idx, row in df.iterrows():
        consulta = row[columna_lim]
        lat, lon = geocodificador(consulta, geolocator, pausa=1)
        if lat is None:
            no_encontradas.append({"Indice": idx, "Original":row.get("Dirección Estandarizada", ""), "Limpia": consulta})
        else:
            df.at[idx, "Latitud"] = lat
            df.at[idx, "Longitud"] = lon

#GUARDAR DOCUMENTO CON COORDENADAS
    encontrados = df[df["Latitud"].notnull() & df["Longitud"].notnull()].copy()
    if not encontrados.empty:
        encontrados.to_excel(doc_coord_exc, index=False)
        centro = [encontrados ["Latitud"].mean(), encontrados["Longitud"].mean()]
        mapa = folium.Map(location=centro, zoom_start=12)
        for _, r in encontrados.iterrows():
            popup_html = f"<b>{r.get('Dirección Estandarizada', '')}</b><br>{r.get('Dirección Estandarizada', '')}"
        folium.CircleMarker(location=[r["Latitud"], r["Longitud"]], radius=5, popup=folium.Popup(popup_html, max_width=300)).add_to(mapa) 
        mapa.save(mapa_int)
        print(f"Mapa guardado en: {mapa_int}")
    else:
        print("No se encontraron coordenadas para crear el mapa.")

#GUARDAR DATOS NO ENCONTRADOS
    if no_encontradas:
        with open ("Direcciones no encontradas.txt", "w", encoding="utf-8") as f:
            for r in no_encontradas:
                f.write(f"{r['Indice']}\t{r['Original']}\t{r['Limpia']}\n")
        print(f"{len(no_encontradas)} direcciones no encontradas guardadas en: 'Direcciones no encontradas.txt'")
    return df, encontrados, no_encontradas

if __name__ == "__main__":
    IN_LIMPIO = "data/Dirección Limpia.xlsx"
    coordenadas(IN_LIMPIO)