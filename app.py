# -*- coding: utf-8 -*-
"""
Created on Fri May 23 23:33:00 2025

@author: Lenovo
"""

import streamlit as st
import pandas as pd
import math
import io

st.title("Rechner für Zylinder und Abstand")

def process(value, length, filter=None):
    [name, x] = value
    possible_n_min = x / (length + 10)
    possible_n_max = x / (length + 2)
    res = [name]
    for i in range(int(math.ceil(possible_n_min)), int(math.ceil(possible_n_max))):
        a1 = (x - i * length) / i
        if (a1 > 2 and a1 < 10 and (not filter or filter(a1))):
            res.append(a1)
    return res

def process_all(values, length, filter=None):
    return list(map(lambda x: process(x, length, filter), values))

def correct_data(data):
    def make_it_lower_than_1000(x):
        [x0, x1] = x
        while (x1 > 1000):
            x1 /= 10
        return [x0, x1]
    return list(map(make_it_lower_than_1000, data))

def check_decimals(num):
    try:
        float_num = float(int(num * 10**5) / 10 ** 5)
        decimal_places = str(float_num).split(".")[1]
        if len(decimal_places) > 2:
            return False
        return True
    except:
        return False

def load_value_from_file(file):
    df = pd.read_excel(file)
    rs = correct_data(df.values.T[0:2].T)
    return rs

uploaded_file = st.file_uploader("Excel-Datei hochladen (xlsx)", type=["xlsx"])

length_input = st.text_input("Etikett Länge (z.B. 50.0)", "0")

if st.button("Berechne und Exportiere"):
    if uploaded_file is None:
        st.error("Bitte zuerst eine Excel-Datei hochladen.")
    else:
        try:
            length = float(length_input.replace(",", "."))
        except ValueError:
            st.error("Bitte eine gültige Zahl für die Etikett Länge eingeben.")
        else:
            values = load_value_from_file(uploaded_file)
            output_full = process_all(values, length)
            filtered_output = process_all(values, length, filter=check_decimals)
            df_full = pd.DataFrame(output_full)
            df_filtered = pd.DataFrame(filtered_output)

            output_buffer = io.BytesIO()
            with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
                df_full.to_excel(writer, index=False, sheet_name='Full Output')
                df_filtered.to_excel(writer, index=False, sheet_name='Gefiltert')

            output_buffer.seek(0)

            st.success("Berechnung abgeschlossen!")

            st.download_button(
                label="Ergebnis herunterladen (Excel)",
                data=output_buffer,
                file_name=f"ergebnis_{length}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
