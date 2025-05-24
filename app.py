# -*- coding: utf-8 -*-
"""
Created on Fri May 23 23:33:00 2025

@author: Lenovo
"""

import streamlit as st
import pandas as pd
import math
import io

st.title("Zylinderabstandsrechner mit dualem Export")

def hat_max_3_dezimalstellen(zahl: float) -> bool:
    """Prüft ob eine Zahl maximal 3 Dezimalstellen hat (ohne Rundung)"""
    num_str = f"{zahl:.10f}"
    if '.' in num_str:
        decimal_part = num_str.split('.')[1].rstrip('0')
        return len(decimal_part) <= 3
    return True

def korrigiere_einheiten(wert: float) -> float:
    """Korrigiert Einheiten für Werte > 1000 (mm zu Meter)"""
    return wert / 1000 if wert > 1000 else wert

def verarbeite_daten(datei) -> list:
    """Liest und bereinigt die Excel-Daten"""
    try:
        df = pd.read_excel(datei)
        df.columns = [col.strip().upper() for col in df.columns]
        
        if not all(col in df.columns for col in ['NAME', 'VALUE']):
            st.error("Erforderliche Spalten 'NAME' oder 'VALUE' nicht gefunden")
            return None
            
        df['VALUE'] = pd.to_numeric(df['VALUE'], errors='coerce')
        df = df.dropna(subset=['VALUE'])
        df['VALUE'] = df['VALUE'].apply(korrigiere_einheiten)
        
        return df[['NAME', 'VALUE']].values.tolist()
        
    except Exception as e:
        st.error(f"Fehler beim Datenimport: {str(e)}")
        return None

def berechne_abstaende(daten: list, laenge: float) -> tuple:
    """Berechnet alle möglichen Abstände und gefilterte Ergebnisse"""
    alle_ergebnisse = []
    gefilterte_ergebnisse = []
    
    for name, messwert in daten:
        zeile_alle = [name]
        zeile_gefiltert = [name]
        
        try:
            n_min = math.ceil(messwert / (laenge + 10))
            n_max = math.ceil(messwert / (laenge + 2))
            
            for n in range(n_min, n_max + 1):
                abstand = (messwert - n * laenge) / n
                if 2 < abstand < 10:
                    abstand_str = f"{abstand:.10f}".rstrip("0").rstrip(".") + " mm"
                    zeile_alle.append(f"{n}x ({abstand_str})")
                    
                    if hat_max_3_dezimalstellen(abstand):
                        zeile_gefiltert.append(f"{n}x ({abstand_str})")
        
        except Exception as e:
            st.warning(f"Fehler bei {name}: {str(e)}")
            continue
        
        alle_ergebnisse.append(zeile_alle)
        if len(zeile_gefiltert) > 1:  # Nur wenn Ergebnisse vorhanden sind
            gefilterte_ergebnisse.append(zeile_gefiltert)
    
    return alle_ergebnisse, gefilterte_ergebnisse

# Streamlit UI
uploaded_file = st.file_uploader("Excel-Datei hochladen", type=["xlsx"])
label_length = st.number_input("Etikettenlänge (mm)", min_value=0.1, value=50.0, step=0.1)

if st.button("Berechnen und Exportieren"):
    if uploaded_file is None:
        st.error("Bitte zuerst eine Excel-Datei hochladen")
    else:
        with st.spinner("Daten werden verarbeitet..."):
            daten = verarbeite_daten(uploaded_file)
            
            if daten:
                alle, gefiltert = berechne_abstaende(daten, label_length)
                
                # Finde maximale Anzahl von Optionen für die Spalten
                max_options_alle = max(len(row) for row in alle) - 1 if alle else 0
                max_options_gefiltert = max(len(row) for row in gefiltert) - 1 if gefiltert else 0
                
                try:
                    # Erstelle DataFrames mit dynamischen Spalten
                    df_alle = pd.DataFrame(alle, columns=["Zylinder"] + [f"Option {i+1}" for i in range(max_options_alle)])
                    df_gefiltert = pd.DataFrame(gefiltert, columns=["Zylinder"] + [f"Option {i+1}" for i in range(max_options_gefiltert)])
                    
                    # Erstelle zwei separate Excel-Dateien
                    output_alle = io.BytesIO()
                    with pd.ExcelWriter(output_alle, engine='openpyxl') as writer:
                        df_alle.to_excel(writer, index=False, sheet_name="Alle Ergebnisse")
                    output_alle.seek(0)
                    
                    output_gefiltert = io.BytesIO()
                    with pd.ExcelWriter(output_gefiltert, engine='openpyxl') as writer:
                        df_gefiltert.to_excel(writer, index=False, sheet_name="Gefilterte Ergebnisse")
                    output_gefiltert.seek(0)
                    
                    # Download-Buttons
                    st.success("Berechnung abgeschlossen!")
                    
                    col1, col2 = st.columns(2)
                    with col1:
                        st.download_button(
                            label="Alle Ergebnisse herunterladen",
                            data=output_alle,
                            file_name=f"alle_abstaende_{label_length}mm.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    
                    with col2:
                        st.download_button(
                            label="Gefilterte Ergebnisse herunterladen",
                            data=output_gefiltert,
                            file_name=f"gefilterte_abstaende_{label_length}mm.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    
                    # Vorschau der Ergebnisse
                    st.subheader("Vorschau der gefilterten Ergebnisse")
                    st.dataframe(df_gefiltert.head(5))
                
                except Exception as e:
                    st.error(f"Fehler beim Erstellen der Excel-Dateien: {str(e)}")
