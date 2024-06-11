# -*- coding: utf-8 -*-
"""
Author: 4wardEnergy Research GmbH
Date: 2024-05-14
Version: 1.0

Reads original consumer data and processes it for further use in lT4dh.

    - Division of the total load into heating load and hot water preparation load
    - depending on the control: "parallel" or "either / or" 
    - storage filling
    - Original data according to consumer_info.csv is compared with existing CSV files in the folder
    - "general_info.csv" contains information about storage control, number of heating circuits, storage sensor status
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

import os
import pandas as pd
import numpy as np
import sys  # for stopping interpreter with sys.exit("Warning-Info")
import datetime as dt
from csv import writer
from csv import reader
from options import *


"""

"""

# -------------------- FILE LOCATIONS -----------------------------------------
pd.set_option('mode.chained_assignment', None)  # disable warnings for channelled expressions
debug_csvs = var_cons_prep.debug_csvs
homedirectory_online = var_cons_prep.home_dir
project_directory_daten_github_online = var_cons_prep.data_dir
project_directory_data_preparation_online = var_cons_prep.raw_data_dir
path_to_file_consumer_info = var_cons_prep.path_to_file_consumer_info
output_directory = var_cons_prep.cons_dir

# -------------------- DEFINITIONS -----------------------------------------

# Define datatypes of input CSV
dtypes_csv = {'Zeitstempel': str,
              'ReglerID': int,
              'akt. Leistung(kW)': float,
              'ges. Waermemenge (kWh)': float,

              'ges. Volumen(m³)': float,
              'Durchfluss (l/h)': float,

              'Vorlauftemperatur (°C)': float,
              'Ruecklauftemperatur (°C)': float,
              'Differenztemperatur(Spreizung) (K)': float,

              'Vorlaufdruck (kPascal)': float,
              'Differenzdruck (kPascal)': float,
              'Differenzdruck (kPascal)': float,

              'Aussentemperatur (°C)': float,
              'WW Speicherfuehler 1 (°C)': float,
              'Speicherfuehler 2 (°C)': float,

              'WW Speicherladepumpe 1': str,
              'Heizkreis 1 Soll (°C)': float,
              'Heizkreis 1 Vorlauffuehler (°C)': float,
              'Heizkreis 1 Raumfuehler (°C)': float,
              'Heizkreis 1 Pumpe': str,

              'Heizkreis 2 Soll (°C)': float,
              'Heizkreis 2 Vorlauffuehler (°C)': float,
              'Heizkreis 2 Raumfuehler (°C)': float,
              'Heizkreis 2 Pumpe': str,

              'Heizkreis 3 Soll (°C)': float,
              'Heizkreis 3 Vorlauffuehler (°C)': float,
              'Heizkreis 3 Raumfuehler (°C)': float,
              'Heizkreis 3 Pumpe': str,

              'Heizkreis 4 Soll (°C)': float,
              'Heizkreis 4 Vorlauffuehler (°C)': float,
              'Heizkreis 4 Raumfuehler (°C)': float,
              'Heizkreis 4 Pumpe': str,

              'letzte Kommunikation (sek)': str,
              'Datenuebertragung (sec)': float}

# -------------------- FUNCTIONS -----------------------------------------


def find_regler(lst):
    """
    Find all files with "Regler" in the name
    """
    result = []
    for s in lst:
        if "Regler" in s:
            result.append(s)
    return result

# Zeitbereich für Daten
cutdate_low = dt.datetime(year=2020, month=11, day=11)
cutdate_high = dt.datetime(year=2024, month=5, day=17)
cutdate_low_2 = dt.datetime(year=2020, month=9, day=13)
cutdate_high_2 = dt.datetime(year=2024, month=12, day=5)

# -------------------- READ INPUT FILES -----------------------------------------

list_to_exclude = ['example_file1.csv', 'exlude_file2.csv']

list_files_in_data_preparation = [file for file in os.listdir(project_directory_data_preparation_online) if file not in list_to_exclude]
list_files_available = find_regler(list_files_in_data_preparation)

csv_consumer_info = pd.read_csv(path_to_file_consumer_info, header=0, on_bad_lines='skip',
                                delimiter=";", decimal=".")

csv_consumer_info['Speicher-Heizkreis'] = csv_consumer_info['Speicher-Heizkreis'].apply(
    lambda x: 'parallel' if str(x).startswith('beides') else x)
csv_consumer_info['Speicher-Heizkreis'] = csv_consumer_info['Speicher-Heizkreis'].apply(
    lambda x: 'entweder/oder' if str(x).startswith('getrennt') else x)

list_regler_names_to_prepare = csv_consumer_info['Regler-ID']
list_regler_names_to_prepare = [str(i) for i in list_regler_names_to_prepare]

list_files_to_prepare = []

for regler_name in list_regler_names_to_prepare:
    # pad regler name to 4 digits to avoid misassociations
    regler_name = regler_name.zfill(4)
    for file_name in list_files_available:
        if regler_name in file_name:
            list_files_to_prepare.append(file_name)


regler_names = [elem.split("_")[1] for elem in list_files_to_prepare]
count_list_files_to_prepare = len(regler_names)

names = regler_names

number_of_storages = count_list_files_to_prepare

general_info_columns = ['Grid number',
                        'technical connected load',
                        'contractual connected load',
                        'connected load',
                        'ww storage sensor',
                        'ww storage state',
                        'storage controller setting',
                        'active heating circuits',
                        'timeseries start with storage',
                        'ww storage controller info']

general_info_index = names[:count_list_files_to_prepare]

general_info = pd.DataFrame(0, columns=general_info_columns, index=general_info_index)

general_info["avg. daily consumption heating winter"] = ""
general_info["avg. daily consumption storage winter"] = ""
general_info["avg. flow temperature winter"] = ""
general_info['avg. return temperature heating winter'] = ""
general_info['avg. return temperature storage winter'] = ""

general_info["avg. daily consumption heating summer"] = ""
general_info["avg. daily consumption storage summer"] = ""
general_info["avg. flow temperature summer"] = ""
general_info['avg. return temperature heating summer'] = ""
general_info['avg. return temperature storage summer'] = ""

general_info["avg. daily consumption heating transition"] = ""
general_info["avg. daily consumption storage transition"] = ""
general_info["avg. flow temperature transition"] = ""
general_info['avg. return temperature heating transition'] = ""
general_info['avg. return temperature storage transition'] = ""

general_info.index.name = 'storages'

factors = [1] * number_of_storages

"""Wenn Speicher gefunden, dann anderes Skript, sonst diesen Teil"""


# read in data
for file in range(count_list_files_to_prepare):

    print()
    # file = 9
    filecount = file + 1
    print('filenumber =', filecount, '/', count_list_files_to_prepare)
    current_controller = int(names[file])
    print('current_controller =', current_controller)
    print('regler_names[file] =', regler_names[file])
    current_row = csv_consumer_info.loc[csv_consumer_info['Regler-ID'] == current_controller, ['Speicher-Heizkreis']].index[0]
    print('current_row =', current_row)
    current_controller_setting = csv_consumer_info.loc[current_row, 'Speicher-Heizkreis']
    print('current_controller_setting =', current_controller_setting)
    current_storage_state = csv_consumer_info.loc[current_row, 'WW Speicher']
    print('current_storage_state =', current_storage_state)
    current_active_heating_circuits = csv_consumer_info.loc[current_row, 'Aktive Heizkreise']
    print('current_active_heating_circuits =', current_active_heating_circuits)

    current_grid_number = csv_consumer_info.loc[current_row, 'Netzplannummer']
    print('current_grid_number =', current_grid_number)
    current_technical_connected_load = csv_consumer_info.loc[current_row, 'Technische Anschlussleistung [kW]']
    print('current_technical_connected_load =', current_technical_connected_load)
    current_contractual_connected_load = csv_consumer_info.loc[current_row, 'Vertragliche Anschlussleistung [kW]']
    print('current_contractual_connected_load =', current_contractual_connected_load)
    current_connected_load = csv_consumer_info.loc[current_row, 'Leistung kW']
    print('current_connected_load =', current_connected_load)

    first_value_heizung_nan = '-'

#
# --- WITHOUT Storage
#     

# Fill in missing column 33 (last data transmission) with default text
    default_text = ';-'
    path_to_file_storage_data = os.path.join(project_directory_data_preparation_online, list_files_to_prepare[file])
    with open(path_to_file_storage_data, 'r') as read_obj, \
            open('XXX.csv', 'w', newline='') as write_obj:
        # Create a csv.reader object from the input file object
        csv_reader = reader(read_obj)
        # Create a csv.writer object from the output file object
        csv_writer = writer(write_obj)
        # Read each row of the input csv file as list
        cntr_1 = 0
        for row in csv_reader:
            if cntr_1 == 0:
                row[0] = 'Zeitstempel;ReglerID;akt. Leistung(kW);ges. Waermemenge (kWh);ges. Volumen(mÂ³);Durchfluss (l/h);Vorlauftemperatur (Â°C);Ruecklauftemperatur (Â°C);Differenztemperatur(Spreizung) (K);Vorlaufdruck (kPascal);Ruecklaufdruck (kPascal);Differenzdruck (kPascal);Aussentemperatur (Â°C);WW Soll (Â°C);WW Speicherfuehler 1 (Â°C);WW Speicherfuehler 2 (Â°C);WW Speicherladepumpe 1;Heizkreis 1 Soll (Â°C);Heizkreis 1 Vorlauffuehler (Â°C);Heizkreis 1 Raumfuehler (Â°C);Heizkreis 1 Pumpe;Heizkreis 2 Soll (Â°C);Heizkreis 2 Vorlauffuehler (Â°C);Heizkreis 2 Raumfuehler (Â°C);Heizkreis 2 Pumpe;Heizkreis 3 Soll (Â°C);Heizkreis 3 Vorlauffuehler (Â°C);Heizkreis 3 Raumfuehler (Â°C);Heizkreis 3 Pumpe;Heizkreis 4 Soll (Â°C);Heizkreis 4 Vorlauffuehler (Â°C);Heizkreis 4 Raumfuehler (Â°C);Heizkreis 4 Pumpe;Datenuebertragung (sec)'
            cntr_1 += 1

            # Append the default text in the row / list
            if row[0].count(";") == 32:
                row[0] = row[0] + default_text
            # Add the updated row / list to the output file
            csv_writer.writerow(row)

    if current_storage_state == 'kein Speicher/keine Info':

        print('NO STORAGE CALCULATION')
        # read in data
        storage_data = pd.read_csv('XXX.csv', sep=';', usecols=[
                                   0, 3, 4, 5, 6, 7, 9, 10, 12, 13, 14, 16, 20, 24, 28, 32, 33], na_values=['-','error'], dtype=dtypes_csv)

        # datachecks
        if bool(storage_data['WW Speicherfuehler 1 (°C)'].isnull().values.all()) is True:
            state_ww_storage_sensor = 'no data'
            print('state_ww_storage_sensor = no data')
        else:
            state_ww_storage_sensor = 'data found'
            print('state_ww_storage_sensor = data found')

        if bool(storage_data['WW Speicherladepumpe 1'].isnull().values.all()) is True:
            ww_storage_controller_info = 'no data'
            print('state_ww_storage_controller = no data')
        else:
            ww_storage_controller_info = 'data found'
            print('state_ww_storage_controller = data found')

        # fehlende Werte im Zeitstempel ergänzen
        # Index vom letzen Wert rausholen
        endind = len(storage_data['Zeitstempel']) - 1
        # Zeitstempel formatieren
        storage_data['Zeitstempel'] = pd.to_datetime(storage_data['Zeitstempel'], format='%d.%m.%Y %H:%M')

        # volle Zeitreihe machen
        idx = pd.date_range(storage_data['Zeitstempel'][0], storage_data['Zeitstempel'][endind], freq='900S')

        # Zeitreihe ergänzen
        storage_data.index = storage_data['Zeitstempel']

        # doppelte Werte löschen (Achtung Zeitumstellung kann zerstört werden)
        storage_data.drop_duplicates(subset='Zeitstempel', keep='last', inplace=True)

        storage_data = storage_data.reindex(idx, fill_value=np.nan)

        # If flow temp. is 0, set all values to nan
        storage_data.loc[storage_data['Vorlauftemperatur (°C)'] == 0] = pd.NA

        # count nan values after reindex and save them to variable nancounts
        nancounts = storage_data['Zeitstempel'].isnull().astype(int).groupby(
            storage_data['Zeitstempel'].notnull().astype(int).cumsum()).cumsum()

        if debug_csvs == 1:
            path_to_file_output_storage_data = os.path.join(
                project_directory_data_preparation_online, regler_names[file] + '_1_interpol.csv')
            storage_data.to_csv(path_to_file_output_storage_data, sep=';', decimal=".", index=False)

        # Zeitstempel richtig ergänzen
        storage_data['Zeitstempel'] = storage_data.index

        # delete useless rows
        storage_data = storage_data[(storage_data['Zeitstempel'] > cutdate_low) & (storage_data['Zeitstempel'] < cutdate_high) |
                                    (storage_data['Zeitstempel'] > cutdate_low_2) & (storage_data['Zeitstempel'] < cutdate_high_2)]
        storage_data.index = [i for i in range(len(storage_data['Zeitstempel']))]

        # determine power
        storage_data['Leistung (kW)'] = 0.977 * 4.182 * (storage_data['Vorlauftemperatur (°C)'] -
                                                         storage_data['Ruecklauftemperatur (°C)']) * storage_data['Durchfluss (l/h)'] / 3600

        if debug_csvs == 1:
            path_to_file_output_storage_data = os.path.join(
                project_directory_data_preparation_online, regler_names[file] + '_2_Leistung.csv')
            storage_data.to_csv(path_to_file_output_storage_data, sep=';', decimal=".", index=False)

        # write nan to power < 0 values
        storage_data['Leistung (kW)'] = storage_data['Leistung (kW)'].mask(storage_data['Leistung (kW)'] < 0)
        storage_data['Leistung (kW)'].fillna(0)

        if debug_csvs == 1:
            path_to_file_output_storage_data = os.path.join(project_directory_data_preparation_online, regler_names[file] + '_23_load.csv')
            storage_data.to_csv(path_to_file_output_storage_data, sep=';', decimal=".", index=False)

        # calculate load for heating and load for hot water heating according to control strategy -> seperate loads
        storage_data['Heizung (kW)'] = storage_data['Leistung (kW)']
        storage_data['Warmwasser (kW)'] = [0] * len(storage_data['Heizung (kW)'])

        if debug_csvs == 1:
            path_to_file_output_storage_data = os.path.join(project_directory_data_preparation_online, regler_names[file] + '_3_load.csv')
            storage_data.to_csv(path_to_file_output_storage_data, sep=';', decimal=".", index=False)

        # write false to nan values
        storage_data['WW Speicherladepumpe 1'].fillna(False, inplace=True)
        storage_data['WW Speicherladepumpe 1'] = storage_data['WW Speicherladepumpe 1'].replace(
            {'True': True, 'true': True, 'False': False, 'false': False})

        if debug_csvs == 1:
            path_to_file_output_storage_data = os.path.join(
                project_directory_data_preparation_online, regler_names[file] + '_4_false_to_nan.csv')
            storage_data.to_csv(path_to_file_output_storage_data, sep=';', decimal=".", index=False)

        storage_data['Heizkreis 1 Pumpe'].fillna(False, inplace=True)
        storage_data['Heizkreis 1 Pumpe'] = storage_data['Heizkreis 1 Pumpe'].replace(
            {'True': True, 'true': True, 'False': False, 'false': False})

        storage_data['Heizkreis 2 Pumpe'].fillna(False, inplace=True)
        storage_data['Heizkreis 2 Pumpe'] = storage_data['Heizkreis 2 Pumpe'].replace(
            {'True': True, 'true': True, 'False': False, 'false': False})

        storage_data['Heizkreis 3 Pumpe'].fillna(False, inplace=True)
        storage_data['Heizkreis 3 Pumpe'] = storage_data['Heizkreis 3 Pumpe'].replace(
            {'True': True, 'true': True, 'False': False, 'false': False})

        storage_data['Heizkreis 4 Pumpe'].fillna(False, inplace=True)
        storage_data['Heizkreis 4 Pumpe'] = storage_data['Heizkreis 4 Pumpe'].replace(
            {'True': True, 'true': True, 'False': False, 'false': False})

        storage_data_filter = storage_data.copy()

        # lowTEMP-specific data preparation
        storage_data = storage_data[['Zeitstempel', 'Leistung (kW)', 'ges. Waermemenge (kWh)',
                                    'ges. Volumen(m³)', 'Durchfluss (l/h)', 'Vorlauftemperatur (°C)', 
                                    'Ruecklauftemperatur (°C)', 'Vorlaufdruck (kPascal)', 'Ruecklaufdruck (kPascal)',
                                    'Datenuebertragung (sec)', 'Aussentemperatur (°C)', 'Heizung (kW)', 'Warmwasser (kW)']]

        storage_data.rename(columns={'Leistung (kW)': 'akt. Leistung(kW)'}, inplace=True)

        # BOOKMARK: Here, the _prepared files are saved for the No Storage case. 
        path_to_file_output_storage_data = os.path.join(output_directory, 'Regler_' + regler_names[file] + '_prepared.csv')
        storage_data.to_csv(path_to_file_output_storage_data, sep=',', decimal=".", index=False)
        print('COMPLETED: No storage found')


#
# --- WITH Storage
#

    elif current_storage_state == 'Speicher gefunden':

        print('STORAGE CALCULATION')
        regelung = current_controller_setting

        storage_data = pd.read_csv('XXX.csv', sep=';', usecols=[
                                   0, 3, 4, 5, 6, 7, 9, 10, 12, 13, 14, 16, 20, 24, 28, 32, 33], na_values=['-', 'error'], dtype=dtypes_csv)

        check_df_storage = storage_data['WW Speicherfuehler 1 (°C)']

        # datachecks
        if bool(storage_data['WW Speicherfuehler 1 (°C)'].isnull().values.all()) is True:
            state_ww_storage_sensor = 'no data'
            print('state_ww_storage_sensor = no data')
        else:
            state_ww_storage_sensor = 'data found'
            print('state_ww_storage_sensor = data found')

        if bool(storage_data['WW Speicherladepumpe 1'].isnull().values.all()) is True:
            ww_storage_controller_info = 'no data'
            print('state_ww_storage_controller = no data')
        else:
            ww_storage_controller_info = 'data found'
            print('state_ww_storage_controller = data found')

        storage_data['Zeitstempel'] = pd.to_datetime(storage_data['Zeitstempel'], format='%d.%m.%Y %H:%M')

        # volle Zeitreihe machen
        endind = len(storage_data['Zeitstempel']) - 1
        idx = pd.date_range(storage_data['Zeitstempel'][0], storage_data['Zeitstempel'][endind], freq='900S')

        # Zeitreihe ergänzen
        storage_data.index = storage_data['Zeitstempel']

        # doppelte Werte löschen (Achtung Zeitumstellung kann zerstört werden)
        storage_data.drop_duplicates(subset='Zeitstempel', keep='last', inplace=True)

        storage_data = storage_data.reindex(idx, fill_value=np.nan)

        # If flow temp. is 0, set all values to nan
        storage_data.loc[storage_data['Vorlauftemperatur (°C)'] == 0] = pd.NA

        # count nan values after reindex and save them to variable nancounts
        nancounts = storage_data['Zeitstempel'].isnull().astype(int).groupby(
            storage_data['Zeitstempel'].notnull().astype(int).cumsum()).cumsum()

        # Zeitstempel richtig ergänzen
        storage_data['Zeitstempel'] = storage_data.index

        # delete useless rows
        storage_data = storage_data[(storage_data['Zeitstempel'] > cutdate_low) & (storage_data['Zeitstempel'] < cutdate_high) |
                                    (storage_data['Zeitstempel'] > cutdate_low_2) & (storage_data['Zeitstempel'] < cutdate_high_2)]
        storage_data.index = [i for i in range(len(storage_data['Zeitstempel']))]

        print(list_files_to_prepare[file], max(storage_data['Durchfluss (l/h)'])/3600)

        # determine power
        storage_data['Leistung (kW)'] = 0.977 * 4.182 * (storage_data['Vorlauftemperatur (°C)'] - storage_data['Ruecklauftemperatur (°C)']) \
            * storage_data['Durchfluss (l/h)'] / 3600

        # write nan to power < 0 values
        storage_data['Leistung (kW)'] = storage_data['Leistung (kW)'].mask(storage_data['Leistung (kW)'] < 0)
        storage_data['Leistung (kW)'].fillna(0)

        # write false to nan values
        storage_data['WW Speicherladepumpe 1'].fillna(False, inplace=True)
        storage_data['WW Speicherladepumpe 1'] = storage_data['WW Speicherladepumpe 1'].replace(
            {'True': True, 'true': True, 'False': False, 'false': False})

        storage_data['Heizkreis 1 Pumpe'].fillna(False, inplace=True)
        storage_data['Heizkreis 1 Pumpe'] = storage_data['Heizkreis 1 Pumpe'].replace(
            {'True': True, 'true': True, 'False': False, 'false': False})

        storage_data['Heizkreis 2 Pumpe'].fillna(False, inplace=True)
        storage_data['Heizkreis 2 Pumpe'] = storage_data['Heizkreis 2 Pumpe'].replace(
            {'True': True, 'true': True, 'False': False, 'false': False})

        storage_data['Heizkreis 3 Pumpe'].fillna(False, inplace=True)
        storage_data['Heizkreis 3 Pumpe'] = storage_data['Heizkreis 3 Pumpe'].replace(
            {'True': True, 'true': True, 'False': False, 'false': False})

        storage_data['Heizkreis 4 Pumpe'].fillna(False, inplace=True)
        storage_data['Heizkreis 4 Pumpe'] = storage_data['Heizkreis 4 Pumpe'].replace(
            {'True': True, 'true': True, 'False': False, 'false': False})

        # calculate load for heating and load for hot water heating according to control strategy -> seperate loads
        if regelung == 'entweder/oder':
            print('STORAGE CALCULATION | entweder/oder')
            storage_data['Warmwasser (kW)'] = storage_data['Leistung (kW)'] * storage_data['WW Speicherladepumpe 1']
            storage_data['Heizung (kW)'] = storage_data['Leistung (kW)'] * (1 - storage_data['WW Speicherladepumpe 1'])

        elif regelung == 'parallel':
            print('STORAGE CALCULATION | parallel')

            copy_leistung = storage_data['Leistung (kW)'].copy(deep=True)
            mask_pumpe = storage_data['WW Speicherladepumpe 1']
            masked_leistung = copy_leistung.mask(mask_pumpe)

            storage_data['Heizung (kW)'] = masked_leistung

            if pd.isnull(storage_data['Heizung (kW)'].iloc[0]):
                storage_data['Heizung (kW)'].iloc[0] = 0
                first_value_heizung_nan = 'first value replaced with 0'
                print('first value was nan')
            else:
                first_value_heizung_nan = 'first value not nan'

            storage_data['Heizung (kW)'].interpolate(inplace=True)

            copy_leistung = storage_data['Leistung (kW)'].copy(deep=True)
            copy_heizung = storage_data['Heizung (kW)'].copy(deep=True)

            storage_data['Warmwasser (kW)'] = copy_leistung - copy_heizung

            mask_set_to_0 = storage_data['Warmwasser (kW)'] < 0
            storage_data.loc[mask_set_to_0, 'Heizung (kW)'] = 0

            storage_data['Warmwasser (kW)'][storage_data['Warmwasser (kW)'] <
                                            0] = storage_data['Leistung (kW)'][storage_data['Warmwasser (kW)'] < 0]

        storage_data_filter = storage_data.copy()
        
        # lowTEMP-specific data preparation
        storage_data = storage_data[['Zeitstempel', 'Leistung (kW)', 'ges. Waermemenge (kWh)',
                                    'ges. Volumen(m³)', 'Durchfluss (l/h)', 'Vorlauftemperatur (°C)', 
                                    'Ruecklauftemperatur (°C)', 'Vorlaufdruck (kPascal)', 'Ruecklaufdruck (kPascal)',
                                    'Datenuebertragung (sec)', 'Aussentemperatur (°C)', 'Heizung (kW)', 'Warmwasser (kW)']]

        storage_data.rename(columns={'Leistung (kW)': 'akt. Leistung(kW)'}, inplace=True)

        # BOOKMARK: Here, the _prepared files are saved for the Storage case. 
        path_to_file_output_storage_data = os.path.join(output_directory, 'Regler_' + regler_names[file] + '_prepared.csv')
        storage_data.to_csv(path_to_file_output_storage_data, sep=',', decimal=".", index=False)

    general_info['Grid number'][regler_names[file]] = current_grid_number
    general_info['technical connected load'][regler_names[file]] = current_technical_connected_load
    general_info['contractual connected load'][regler_names[file]] = current_contractual_connected_load
    general_info['connected load'][regler_names[file]] = current_connected_load

    general_info['ww storage sensor'][regler_names[file]] = state_ww_storage_sensor
    general_info['ww storage state'][regler_names[file]] = current_storage_state
    general_info['storage controller setting'][regler_names[file]] = current_controller_setting
    general_info['active heating circuits'][regler_names[file]] = current_active_heating_circuits
    general_info['timeseries start with storage'][regler_names[file]] = first_value_heizung_nan
    general_info['ww storage controller info'][regler_names[file]] = ww_storage_controller_info

#
# --- Computation of seasonal characteristics
#

    filtered_data_storage = storage_data_filter.copy()
    filtered_data_storage = filtered_data_storage.fillna(0)

    # filter for winter months
    filtered_data_winter = filtered_data_storage[filtered_data_storage['Zeitstempel'].dt.month.isin([11, 12, 1, 2])]
    filtered_data_winter.set_index('Zeitstempel', inplace=True)

    filtered_data_winter_storage_info_mask = filtered_data_winter[filtered_data_winter['WW Speicherladepumpe 1'] == True]
    filtered_data_winter_heating_info_mask = filtered_data_winter[(filtered_data_winter['Heizkreis 1 Pumpe'] == True) | (
        filtered_data_winter['Heizkreis 2 Pumpe'] == True) | (filtered_data_winter['Heizkreis 3 Pumpe'] == True) | (filtered_data_winter['Heizkreis 4 Pumpe'] == True)]
    filtered_data_winter = filtered_data_winter.drop(
        columns=['Heizkreis 1 Pumpe', 'Heizkreis 2 Pumpe', 'Heizkreis 3 Pumpe', 'Heizkreis 4 Pumpe', 'WW Speicherladepumpe 1'])
    filtered_data_daily_avg_winter = filtered_data_winter.resample('D').mean()
    filtered_data_seasonal_avg_winter = filtered_data_daily_avg_winter.mean()
    # filter for summer months
    filtered_data_summer = filtered_data_storage[filtered_data_storage['Zeitstempel'].dt.month.isin([5, 6, 7, 8])]
    filtered_data_summer.set_index('Zeitstempel', inplace=True)
    filtered_data_summer_storage_info_mask = filtered_data_summer[filtered_data_summer['WW Speicherladepumpe 1'] == True]
    filtered_data_summer_heating_info_mask = filtered_data_summer[(filtered_data_summer['Heizkreis 1 Pumpe'] == True) | (
        filtered_data_summer['Heizkreis 2 Pumpe'] == True) | (filtered_data_summer['Heizkreis 3 Pumpe'] == True) | (filtered_data_summer['Heizkreis 4 Pumpe'] == True)]
    filtered_data_summer = filtered_data_summer.drop(
        columns=['Heizkreis 1 Pumpe', 'Heizkreis 2 Pumpe', 'Heizkreis 3 Pumpe', 'Heizkreis 4 Pumpe', 'WW Speicherladepumpe 1'])
    filtered_data_daily_avg_summer = filtered_data_summer.resample('D').mean()
    filtered_data_seasonal_avg_summer = filtered_data_daily_avg_summer.mean()
    # filter for transition months
    filtered_data_transition = filtered_data_storage[filtered_data_storage['Zeitstempel'].dt.month.isin([3, 4, 9, 10])]
    filtered_data_transition.set_index('Zeitstempel', inplace=True)
    filtered_data_transition_storage_info_mask = filtered_data_transition[filtered_data_transition['WW Speicherladepumpe 1'] == True]
    filtered_data_transition_heating_info_mask = filtered_data_transition[(filtered_data_transition['Heizkreis 1 Pumpe'] == True) | (
        filtered_data_transition['Heizkreis 2 Pumpe'] == True) | (filtered_data_transition['Heizkreis 3 Pumpe'] == True) | (filtered_data_transition['Heizkreis 4 Pumpe'] == True)]
    filtered_data_transition = filtered_data_transition.drop(
        columns=['Heizkreis 1 Pumpe', 'Heizkreis 2 Pumpe', 'Heizkreis 3 Pumpe', 'Heizkreis 4 Pumpe', 'WW Speicherladepumpe 1'])
    filtered_data_daily_avg_transition = filtered_data_transition.resample('D').mean()
    filtered_data_seasonal_avg_transition = filtered_data_daily_avg_transition.mean()

    filtered_data_return_storage_winter = filtered_data_winter_storage_info_mask.drop(
        columns=['WW Speicherladepumpe 1', 'Heizkreis 1 Pumpe', 'Heizkreis 2 Pumpe', 'Heizkreis 3 Pumpe', 'Heizkreis 4 Pumpe'])
    filtered_data_return_storage_seasonal_avg_winter = filtered_data_return_storage_winter['Ruecklauftemperatur (°C)'].mean()
    filtered_data_return_heating_winter = filtered_data_winter_heating_info_mask.drop(
        columns=['WW Speicherladepumpe 1', 'Heizkreis 1 Pumpe', 'Heizkreis 2 Pumpe', 'Heizkreis 3 Pumpe', 'Heizkreis 4 Pumpe'])
    filtered_data_return_heating_seasonal_avg_winter = filtered_data_return_heating_winter['Ruecklauftemperatur (°C)'].mean()

    filtered_data_return_storage_summer = filtered_data_summer_storage_info_mask.drop(
        columns=['WW Speicherladepumpe 1', 'Heizkreis 1 Pumpe', 'Heizkreis 2 Pumpe', 'Heizkreis 3 Pumpe', 'Heizkreis 4 Pumpe'])
    filtered_data_return_storage_seasonal_avg_summer = filtered_data_return_storage_summer['Ruecklauftemperatur (°C)'].mean()
    filtered_data_return_heating_summer = filtered_data_summer_heating_info_mask.drop(
        columns=['WW Speicherladepumpe 1', 'Heizkreis 1 Pumpe', 'Heizkreis 2 Pumpe', 'Heizkreis 3 Pumpe', 'Heizkreis 4 Pumpe'])
    filtered_data_return_heating_seasonal_avg_summer = filtered_data_return_heating_summer['Ruecklauftemperatur (°C)'].mean()

    filtered_data_return_storage_transition = filtered_data_transition_storage_info_mask.drop(
        columns=['WW Speicherladepumpe 1', 'Heizkreis 1 Pumpe', 'Heizkreis 2 Pumpe', 'Heizkreis 3 Pumpe', 'Heizkreis 4 Pumpe'])
    filtered_data_return_storage_seasonal_avg_transition = filtered_data_return_storage_transition['Ruecklauftemperatur (°C)'].mean()
    filtered_data_return_heating_transition = filtered_data_transition_heating_info_mask.drop(
        columns=['WW Speicherladepumpe 1', 'Heizkreis 1 Pumpe', 'Heizkreis 2 Pumpe', 'Heizkreis 3 Pumpe', 'Heizkreis 4 Pumpe'])
    filtered_data_return_heating_seasonal_avg_transition = filtered_data_return_heating_transition['Ruecklauftemperatur (°C)'].mean()

    general_info['avg. daily consumption heating summer'][regler_names[file]] = filtered_data_seasonal_avg_summer['Heizung (kW)']
    general_info['avg. daily consumption heating winter'][regler_names[file]] = filtered_data_seasonal_avg_winter['Heizung (kW)']
    general_info['avg. daily consumption heating transition'][regler_names[file]] = filtered_data_seasonal_avg_transition['Heizung (kW)']

    general_info['avg. daily consumption storage summer'][regler_names[file]] = filtered_data_seasonal_avg_summer['Warmwasser (kW)']
    general_info['avg. daily consumption storage winter'][regler_names[file]] = filtered_data_seasonal_avg_winter['Warmwasser (kW)']
    general_info['avg. daily consumption storage transition'][regler_names[file]] = filtered_data_seasonal_avg_transition['Warmwasser (kW)']

    general_info['avg. flow temperature summer'][regler_names[file]] = filtered_data_seasonal_avg_summer['Vorlauftemperatur (°C)']
    general_info['avg. flow temperature winter'][regler_names[file]] = filtered_data_seasonal_avg_winter['Vorlauftemperatur (°C)']
    general_info['avg. flow temperature transition'][regler_names[file]] = filtered_data_seasonal_avg_transition['Vorlauftemperatur (°C)']

    general_info['avg. return temperature heating summer'][regler_names[file]] = filtered_data_return_heating_seasonal_avg_summer
    general_info['avg. return temperature heating winter'][regler_names[file]] = filtered_data_return_heating_seasonal_avg_winter
    general_info['avg. return temperature heating transition'][regler_names[file]] = filtered_data_return_heating_seasonal_avg_transition

    general_info['avg. return temperature storage summer'][regler_names[file]] = filtered_data_return_storage_seasonal_avg_summer
    general_info['avg. return temperature storage winter'][regler_names[file]] = filtered_data_return_storage_seasonal_avg_winter
    general_info['avg. return temperature storage transition'][regler_names[file]] = filtered_data_return_storage_seasonal_avg_transition

    print('regler_names[file] =', regler_names[file])
    print('filtered_data_seasonal_avg_summer =', filtered_data_seasonal_avg_summer['Ruecklauftemperatur (°C)'])
    print('filtered_data_seasonal_avg_winter =', filtered_data_seasonal_avg_winter['Ruecklauftemperatur (°C)'])
    print('filtered_data_seasonal_avg_transition =', filtered_data_seasonal_avg_transition['Ruecklauftemperatur (°C)'])


path_to_file_output_general_info = os.path.join(output_directory, 'general_info.csv')
general_info.to_csv(path_to_file_output_general_info, sep=';', decimal=".", index=True)
