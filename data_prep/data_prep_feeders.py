# -*- coding: utf-8 -*-
"""
Author: 4wardEnergy Research GmbH
Date: 2024-05-14
Version: 1.0

Reads original feeder data and processes it for further use in lT4dh. 
It extracts customer numbers from the topology file, fills them with controller IDs 
from the topology file, and processes time series controller CSVs 
(only those whose headers start with "Timestamp"). The program checks if sum(WW_soll) != 0, 
creates a column indicating if storage is found, counts active heating circuits, and 
checks if heating circuits and storage pump are on simultaneously, outputting YES if 
the threshold (8/100 timesteps) is exceeded.
"""
"""
Copyright (C) 2024  4wardEnergy Research GmbH

This program is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details.

You should have received a copy of the GNU General Public License
along with this program.  If not, see <https://www.gnu.org/licenses/>.
"""

import pandas as pd
import numpy as np
import os
from csv import writer
from csv import reader
from options import *

# ORDNERPFAD MIT DEN ROHDATEN
home_directory = os.path.expanduser( '~' )
path = var_cons_prep.raw_data_dir
os.chdir(path)
os.getcwd()
files = os.listdir()
default_text = ';a'

# ORDNERPFAD FÜR OUTPUT
out_path = var_cons_prep.cons_dir

# NUR EINSPEISER WÄHLEN
filtered_files = []
for i, file in enumerate(files):
    if ".csv" in file:
        if "Router_WMZ" in file:
            filtered_files.append(file)

files = filtered_files

# ID DES REGLERS
controller_id = [ elem[11] for elem in files ]

# SCHLEIFE DURCHLAEUFT ALLE DATENSAETZE
for cntr_0 in range(len(files)):
    with open(files[cntr_0], 'r') as read_obj, \
            open('XXX.csv', 'w', newline='') as write_obj:
        # Create a csv.reader object from the input file object
        csv_reader = reader(read_obj)
        # Create a csv.writer object from the output file object
        csv_writer = writer(write_obj)
        # Read each row of the input csv file as list
        cntr_1 = 0
        for row in csv_reader:
            if cntr_1 == 0:
                row[0] = 'Zeitstempel;ReglerID;akt. Leistung(kW);ges. Waermemenge (kWh);ges. Volumen(m^3);Durchfluss (l/h);Vorlauftemperatur (Â°C);Ruecklauftemperatur (Â°C);Differenztemperatur(Spreizung) (K);Vorlaufdruck (kPascal);Ruecklaufdruck (kPascal);Differenzdruck (kPascal);WW Soll (Â°C);WW Speicherfuehler 1 (Â°C);WW Speicherfuehler 2 (Â°C);WW Speicherladepumpe 1;Heizkreis 1 Soll (Â°C);Heizkreis 1 Vorlauffuehler (Â°C);Heizkreis 1 Raumfuehler (Â°C);Heizkreis 1 Pumpe;Heizkreis 2 Soll (Â°C);Heizkreis 2 Vorlauffuehler (Â°C);Heizkreis 2 Raumfuehler (Â°C);Heizkreis 2 Pumpe;Heizkreis 3 Soll (Â°C);Heizkreis 3 Vorlauffuehler (Â°C);Heizkreis 3 Raumfuehler (Â°C);Heizkreis 3 Pumpe;Heizkreis 4 Soll (Â°C);Heizkreis 4 Vorlauffuehler (Â°C);Heizkreis 4 Raumfuehler (Â°C);Heizkreis 4 Pumpe;Datenuebertragung (sec)'
            cntr_1 += 1

            # Append the default text in the row / list
            if row[0].count(";") == 31:
                #row[0].append(default_text)
                row[0] = row[0] + ";-"
            # Add the updated row / list to the output file
            csv_writer.writerow(row)
    try:
        #######################################################################
        # DATEN EINLESEN ######################################################
        #######################################################################
        #storage_data = pd.read_csv('XXX.csv', sep = ';' ,usecols = [0, 2, 3, 4, 5, 6, 7, 9, 10, 32], na_values = '-', low_memory=False)
        #print(output.decode(encoding="utf-8", errors="replace").split('\n'))
        print(f"Reading data: {files[cntr_0]}")
        storage_data = pd.read_csv('XXX.csv', sep = ';' ,usecols = [0, 2, 3, 4, 5, 6, 7, 9, 10, 32], na_values = '-', low_memory = False)

        # Spalte 0...Zeit [TT.MM.JJJJ hh:mm]
        # Spalte 2...Aktuelle Leistung [kW]
        # Spalte 3...Gesamte Wärmemenge [kWh]
        # Spalte 4...Gesamter Volumenstrom [m³]
        # Spalte 5...Durchfluss [l/h]
        # Spalte 6...Vorlauftemperatur [°C]
        # Spalte 7...Rücklauftemperatur [°C]
        # Spalte 9...Vorlaufdruck [kPa]
        # Spalte 10..Rücklaufdruck [kPa]
        # Spalte 32..Verstrichene Zeit seit der letzten Datenübermittlung [s]

        # INDEX VOM LETZTEN WERT
        endind = len(storage_data['Zeitstempel'])-1

        # ZEITSTEMPEL FORMATIEREN
        storage_data['Zeitstempel'] = pd.to_datetime(storage_data['Zeitstempel'], format='%d.%m.%Y %H:%M')

        # ZEITSTEMPEL FUELLEN
        idx = pd.date_range(storage_data['Zeitstempel'][0], storage_data['Zeitstempel'][endind], freq = '900S' )

        # ZEITREIHE ERGAENZEN
        storage_data.index = storage_data['Zeitstempel']

        # DOPPELTE WERTE LOESCHEN (Achtung Zeitumstellung kann zerstört werden)
        storage_data.drop_duplicates( subset = 'Zeitstempel', keep = 'last', inplace = True )
        
        storage_data = storage_data.reindex( idx, fill_value = np.nan )

        # ZEITSTEMPEL RICHTIG ERGAENZEN
        storage_data['Zeitstempel'] = storage_data.index

        # ZEITBEREICH FESTLEGEN
        #storage_data = storage_data[((storage_data['Zeitstempel'] >= dt.datetime(year = 2021, month = 5, day = 2, hour = 0)) & (storage_data['Zeitstempel'] <= dt.datetime(year = 2021, month = 6, day = 15, hour = 0))]

        # UEBERFLUESSIGE SPALTEN LOESCHEN
        storage_data.drop( ['Zeitstempel'], inplace = True, axis = 1 )

        # WENN VORLAUFTEMPERATUR == 0: ALLE WERTE NaN SETZEN
        storage_data.loc[storage_data['Vorlauftemperatur (°C)'] == 0] = pd.NA

        # SPEICHERN
        storage_data.to_csv(os.path.join(out_path, 'Router_WMZ_'+controller_id[cntr_0] + '_prepared.csv'))
    except:
        print("Konvertierung abgebrochen!")
        print(str(controller_id[cntr_0]))
os.remove('XXX.csv')

