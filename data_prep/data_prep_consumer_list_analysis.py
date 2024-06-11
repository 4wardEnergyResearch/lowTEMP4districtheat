# -*- coding: utf-8 -*-
"""
Author: 4wardEnergy Research GmbH
Date: 2024-05-14
Version: 1.0

Takes the customer number from the network topology and fills it with controller IDs from the network topology.
Scans controller CSVs from time series (only those whose header begins with "Timestamp") and checks if the sum(WW_soll) != 0.
Creates a column with information on whether storage was found.
Counts active heating circuits.
Checks if heating circuits and storage pump are on simultaneously. Outputs YES if the threshold is exceeded (8/100 time steps).
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
import os.path
import time
import sys  # for stopping the interpreter with sys.exit()
from datetime import datetime
from options import *
from simulation.Auxiliary_functions import *


# -------------------- PARAMETERS -----------------------------------------

threshold_active_heizkreis = var_cons_list_analysis.threshold_active_heizkreis

# -------------------- FILE LOCATIONS -----------------------------------------

homedirectory_online = var_cons_list_analysis.project_dir
project_directory_online = var_cons_list_analysis.data_dir
target_path_netztopologie = var_cons_list_analysis.topology_file
target_directory_zeitreihen = var_cons_list_analysis.raw_data_dir
output_directory = target_directory_zeitreihen



# -------------------- DEFINITIONS -------------------------------------------

# Define column labels of Speicherpumpe and Heizkreise
list_columns_controller = []
list_columns_controller.append('Heizkreis 1 Pumpe')
list_columns_controller.append('Heizkreis 2 Pumpe')
list_columns_controller.append('Heizkreis 3 Pumpe')
list_columns_controller.append('Heizkreis 4 Pumpe')
list_columns_controller_speicher = ['WW Speicherladepumpe 1']

# Define datatypes of input CSV
dtypes_csv = {'Zeitstempel': str,
              'ReglerID': int,
              'akt. Leistung(kW)': float,
              'ges. Waermemenge (kWh)': float,
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

              'letzte Kommunikation (sek)': str}


# -------------------- FUNCTIONS -----------------------------------------

def find_regler(lst):
    """
    Find all files containing "Regler" in their name.
    """
    result = []
    for s in lst:
        if "Regler" in s:
            result.append(s)
    return result


def remove_items(list1, list2):
    """
    Remove items from list1 that are in list2.
    """
    return [x for x in list1 if x not in list2]


def get_columns_with_true_value(df):
    """ 
    Get all columns with "true" in the dataframe.
    """
    for column in df.columns:
        if df[column].str.contains("true", case=False).any():
            list_heizkreise_active.append(column)
    return list_heizkreise_active


def count_true_occurrences(df, x):
    """ 
    Count the number of True values in each column and check if any column has more than x True values.
    """
    # Count the number of True values in each column
    true_counts = df.apply(lambda col: col.eq('true').sum())
    total_timesteps_count = len(df) + 0.0001  # avoid DIV/0

    # Check if any column has more than x True values
    if (true_counts/total_timesteps_count > x).any():
        return 1
    else:
        return 0


# -------------------- START -----------------------------------------


timer_script_start = time.time()
start_time = time.time()
current_time = datetime.now()
formatted_time = current_time.strftime("%Y%m%d_%H%M")


# -------------------- LOAD FILES -----------------------------------------

list_to_exclude = ['exlude_file1.csv', 'exlude_file2.csv']

list_files_zeitreihen = [file for file in os.listdir(target_directory_zeitreihen) if file not in list_to_exclude]
list_files_consumer = find_regler(list_files_zeitreihen)

# Read the network topology file into a pandas dataframe
df_netztopologie = pd.read_excel(target_path_netztopologie, dtype=str)

# Delete all rows where "aktiv" is not "ja"
df_netztopologie = df_netztopologie[df_netztopologie['aktiv'] == 'ja']
# Select only the columns "Netzplan ID", "Abnehmer" and "Nennleistung [kW]"
df_netztopologie = df_netztopologie[['Netzplan ID', 'Abnehmer', 'Nennleistung [kW]']]
# Delete all rows where "Abnehmer" is empty
df_netztopologie = df_netztopologie[df_netztopologie['Abnehmer'].notna()]
# Reset index
df_netztopologie = df_netztopologie.reset_index(drop=True)
# Reorder and rename columns to fit needed format
df_consumer_info = df_netztopologie.copy(deep=True)
# Set dtype of column"Nennleistung [kW]" to float
df_consumer_info['Nennleistung [kW]'] = df_consumer_info['Nennleistung [kW]'].astype(float)
df_consumer_info['Vertragliche Anschlussleistung [kW]'] = df_consumer_info['Nennleistung [kW]']
# Reorder columns
df_consumer_info = df_consumer_info[['Netzplan ID', 'Nennleistung [kW]', 'Vertragliche Anschlussleistung [kW]', 'Abnehmer']]
# Rename columns
df_consumer_info = df_consumer_info.rename(columns={'Netzplan ID': 'Netzplannummer', 
                                                    'Nennleistung [kW]': 'Technische Anschlussleistung [kW]',
                                                    'Abnehmer': 'Regler-ID'})
# New column "Leistung kW"
df_consumer_info['Leistung kW'] = df_consumer_info['Technische Anschlussleistung [kW]']


# -------------------- FILL LEGACY COLUMNS -----------------------------------------

df_consumer_info['WW Speicher'] = ""
df_consumer_info['Speicher-Heizkreis'] = ""
df_consumer_info['Aktive Heizkreise'] = ""
df_consumer_info['Speicherbeladungen'] = ""
df_consumer_info['Fehlermeldungen'] = ""
df_consumer_info['Info'] = ""

# check for missing headers and add files to ignore list
list_files_consumer_defective = []
list_regler_not_in_list = []
log = []
log.append(f'Time: {formatted_time}')

for i, value in enumerate(list_files_consumer):
    path_to_current_consumer = os.path.join(target_directory_zeitreihen, list_files_consumer[i])
    check_missing_headers = pd.read_csv(path_to_current_consumer, header=0, nrows=5, on_bad_lines='skip', delimiter=";", decimal=".")

    if check_missing_headers.columns[0] != "Zeitstempel":
        list_files_consumer_defective.append(value)

list_files_consumer_working = remove_items(list_files_consumer, list_files_consumer_defective)

# -------------------- LOOP OVER ALL FILES IN DIRECTORY CONTAINING "Regler_####" -----------------------------------------

length_list_files_consumer_working = len(list_files_consumer_working)
# length_list_files_consumer_working = 1  # run either full list or less for testing

for i in range(length_list_files_consumer_working):

    path_to_current_consumer = os.path.join(target_directory_zeitreihen, list_files_consumer[i])

    print("i =", i)
    log.append(f'i = {i}')

    # Extract current Regler-ID
    current_consumer_ID = list_files_consumer[i].split('_')[1]
    log.append(f'current_consumer_ID = {current_consumer_ID}')

    current_consumer_zeitreihe = pd.read_csv(path_to_current_consumer, parse_dates=['Zeitstempel'],
                                             header=0, on_bad_lines='skip', delimiter=";", decimal=".",
                                             dtype=dtypes_csv, na_values={'-', 'error'}, dayfirst=True)

    current_consumer_zeitreihe['Zeitstempel'] = pd.to_datetime(current_consumer_zeitreihe['Zeitstempel'], format="%d.%m.%Y %H:%M")
    current_consumer_zeitreihe['WW Soll (°C)'] = current_consumer_zeitreihe['WW Soll (°C)'].replace('-', 0)
    current_consumer_zeitreihe['WW Soll (°C)'] = current_consumer_zeitreihe['WW Soll (°C)'].astype(float)

    # -------------------- CHECK FOR WARM WATER STORAGE -----------------------------------------

    if current_consumer_zeitreihe['WW Soll (°C)'].sum() == 0:
        try:
            current_row = df_consumer_info.loc[df_consumer_info['Regler-ID'] == current_consumer_ID].index[0]
        except:
            print('consumer', current_consumer_ID, ' not in list')
            log.append(f'current_consumer_ID {current_consumer_ID} not in list')
        else:
            current_row = df_consumer_info.loc[df_consumer_info['Regler-ID'] == current_consumer_ID].index[0]
            df_consumer_info.at[current_row, 'WW Speicher'] = 'kein Speicher/keine Info'
            log.append(f'current_row = {current_row}')
            log.append(f'{current_consumer_ID}: kein Speicher/keine Info')
    else:
        try:
            current_row = df_consumer_info.loc[df_consumer_info['Regler-ID'] == current_consumer_ID].index[0]
        except:
            print('consumer', current_consumer_ID, ' not in list')
            log.append(f'current_consumer_ID {current_consumer_ID} not in list')

        else:
            current_row = df_consumer_info.loc[df_consumer_info['Regler-ID'] == current_consumer_ID].index[0]
            df_consumer_info.at[current_row, 'WW Speicher'] = 'Speicher gefunden'
            log.append(f'current_row = {current_row}')
            log.append(f'{current_consumer_ID}: Speicher gefunden')

    # -------------------- CHECK FOR SEPARATE FEEDING OF SPEICHER AND HEIZKREIS ---------------------------------
    # Check if Speicher-Pumpe is "true" and either of Heizkreis 1-4 is also "true"

    columns_controller = current_consumer_zeitreihe[list_columns_controller_speicher]
    columns_controller = columns_controller.join(current_consumer_zeitreihe[list_columns_controller])
    columns_controller = columns_controller.astype(str)
    columns_controller = columns_controller.replace({'True': 'true'})
    columns_controller = columns_controller.replace({'False': 'false'})

    mask_true = columns_controller.iloc[:, 0] == "true"
    df_controller_speicher_true = columns_controller[mask_true]

    df_controller_speicher_true = df_controller_speicher_true.drop(list_columns_controller_speicher, axis=1)

    count_ww_storage_pump_active = len(df_controller_speicher_true)
    log.append(f'count_ww_storage_pump_active = {count_ww_storage_pump_active}')

    count_true = df_controller_speicher_true.eq("true").sum()
    count_true_total = count_true.sum()

    check_true_counts = df_controller_speicher_true.apply(lambda col: col.eq('true').sum())
    check_total_timesteps_count = len(df_controller_speicher_true) + 0.0001

    log.append(f'check_true_counts = {check_true_counts}')
    log.append(f'check_total_timesteps_count = {check_total_timesteps_count}')

    # if df_controller_speicher_true.apply(lambda x: x.str.contains('true')).any().any():
    if count_true_occurrences(df_controller_speicher_true, threshold_active_heizkreis) == 1:

        try:
            current_row = df_consumer_info.loc[df_consumer_info['Regler-ID'] == current_consumer_ID].index[0]
        except:
            list_regler_not_in_list.append(current_consumer_ID)
            print(current_consumer_ID, 'added to list_regler_not_in_list')

            log.append(f'current_consumer_ID = {current_consumer_ID}')
            log.append(f'{current_consumer_ID} added to list_regler_not_in_list')
        else:
            current_row = df_consumer_info.loc[df_consumer_info['Regler-ID'] == current_consumer_ID].index[0]
            df_consumer_info.at[current_row, 'Speicher-Heizkreis'] = f"beides in {str(count_true_total)} timesteps"
            print('Speicher und Heizkreis GEMEINSAM')

            list_heizkreise_active = []
            list_heizkreise_active = get_columns_with_true_value(df_controller_speicher_true)
            count_heizkreise_active = len(list_heizkreise_active)
            df_consumer_info.at[current_row, 'Aktive Heizkreise'] = count_heizkreise_active
            print(f'{count_heizkreise_active} Heizkreis(e) aktiv')

            log.append('Speicher und Heizkreis GEMEINSAM')
            log.append(f'current_row = {current_row}')
            log.append(f'count_heizkreise_active = {count_heizkreise_active}')

            df_consumer_info.at[current_row, 'Speicherbeladungen'] = count_ww_storage_pump_active

            log.append(f'current_row = {current_row}')
            log.append(f'Speicherbeladungen = {count_ww_storage_pump_active}')

    else:
        try:
            current_row = df_consumer_info.loc[df_consumer_info['Regler-ID'] == current_consumer_ID].index[0]
        except:
            list_regler_not_in_list.append(current_consumer_ID)
            print(current_consumer_ID, 'added to list_regler_not_in_list')

            log.append(f'current_consumer_ID = {current_consumer_ID}')
            log.append(f'{current_consumer_ID} added to list_regler_not_in_list')
            current_row = 0

        else:
            current_row = df_consumer_info.loc[df_consumer_info['Regler-ID'] == current_consumer_ID].index[0]
            df_consumer_info.at[current_row, 'Speicher-Heizkreis'] = 'getrennt'
            print('Speicher und Heizkreis GETRENNT')

            list_heizkreise_active = []
            list_heizkreise_active = get_columns_with_true_value(columns_controller.drop(0))
            count_heizkreise_active = len(list_heizkreise_active)
            df_consumer_info.at[current_row, 'Aktive Heizkreise'] = count_heizkreise_active
            print(f' {count_heizkreise_active} Heizkreis(e) aktiv')

            log.append('Speicher und Heizkreis GETRENNT')
            log.append(f'current_row = {current_row}')
            log.append(f'count_heizkreise_active = {count_heizkreise_active}')

    log.append('-')
    print('current_row', current_row)
    print('current_consumer_ID', current_consumer_ID)
    # print('------- dtype', current_consumer_zeitreihe.dtypes)
    print('------------------------')

df_consumer_info.at[0, 'Fehlermeldungen'] = f"Files which DID NOT work: {str(list_files_consumer_defective)}"
df_consumer_info.at[1, 'Fehlermeldungen'] = f"Regler as csv but NOT in Abnehmerliste: {str(list_regler_not_in_list)}"

# -------------------- OPTIONAL OUTPUT -----------------------------------------

# df_consumer_info.at[0, 'Info'] = f"WW beladungen: {str(count_ww_storage_pump_active)}"
# df_consumer_info.at[1, 'Info'] = f"Aktive Heizkreise: {str(count_heizkreise_active)}"

print('')
print('------------------------')
print('SUMMARY:')
print('------------------------')

print('Files which DID NOT work: \n', list_files_consumer_defective)
print('------------------------')

print('Regler NOT in consumer_info list: \n', list_regler_not_in_list)
print('------------------------')

# -------------------- EXPORT -----------------------------------------
print('EXPORT Info: \n')

if not os.path.exists(output_directory):
    os.makedirs(output_directory)
    print(f"Folder '{output_directory}' created successfully! \n")
else:
    print(f"Folder '{output_directory}' already exists. \n")

export_filename_df_consumer_info = 'consumer_info.csv'

df_consumer_info.to_csv(os.path.join(output_directory, export_filename_df_consumer_info), sep=';', decimal=".", index=False)
print('CSV:', export_filename_df_consumer_info, 'created.')
print('export path:', output_directory, '\n')

logfile_name = formatted_time + '_log.txt'
logfile_path = os.path.join(output_directory, logfile_name)

with open(logfile_path, "w") as file:
    for item in log:
        file.write(str(item) + "\n")
print('LOG:', logfile_name, 'created.')
print('export path:', output_directory, '\n')

print_red(f"The raw data from the following consumers is defective: {list_files_consumer_defective}")

timer_script_end = time.time()
print('Finished script in', round(timer_script_end - timer_script_start, 0), 's.')
