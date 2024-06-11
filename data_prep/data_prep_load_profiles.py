# -*- coding: utf-8 -*-
"""
Author: 4wardEnergy Research GmbH
Date: 2024-05-14
Version: 1.0

Processes load profiles for different building types based on historical
data, building characteristics, and standardized profiles. It calculates customized load
profiles for heating and sanitary hot water (SHW) consumption, estimates yearly loads, and
saves the results for further use in the simulation. The program also reads the topology and general
information files, handles missing data, and generates diagnostics for data quality.
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

# IMPORTS #####################################################################
###############################################################################
###############################################################################
###############################################################################
import pandas as pd
import os
import termcolor
import openpyxl
from pathlib import Path
import seaborn as sns
import matplotlib.pyplot as plt
import numpy as np
import importlib

# Disable pandas warnings
pd.options.mode.chained_assignment = None

# Import options
from options import *

# FUNCTION DEFINITIONS ########################################################
###############################################################################
###############################################################################
###############################################################################
def print_red(text):
    print(termcolor.colored(text, 'red'))

def print_yellow(text):
    print(termcolor.colored(text, 'yellow'))

def print_green(text):
    print(termcolor.colored(text, 'green'))

def print_blue(text):
    print(termcolor.colored(text, 'blue'))

def get_scale_factor(hourly_profile: pd.DataFrame, yearly_load: float) -> float:
    """Calculates the scale factor for the hourly profile given the
    yearly load by normalizing the integral over the yearly profile
    to the yearly load.

    :param hourly_profile: Hotmaps profile for a year with hourly resolution
    :type hourly_profile: pd.DataFrame
    :param yearly_load: Yearly load of the consumer in kWh
    :type yearly_load: float
    :return: The scale factor that must be applied to the profile
    :rtype: float
    """
    sum_loads = hourly_profile["power [kW]"].sum()
    scale_factor = yearly_load/sum_loads
    return scale_factor

def load_load_profile_heating(var_load_profiles, building_type: str) -> pd.DataFrame:
    """Loads the heating load profile for the given building type.

    :param building_type: Indicates the type of building for which the 
                          load profile should be loaded. See match statement.
    :type building_type: str
    :return: The load profile.
    :rtype: pd.DataFrame
    """    
    match building_type:
        case "ind:papier":
            load_profile_file_name = "hotmaps_task_2.7_load_profile_industry_paper_generic.csv"
        case "ind:leb":
            load_profile_file_name = "hotmaps_task_2.7_load_profile_industry_food_and_tobacco_generic.csv"
        case "ind:stahl":
            load_profile_file_name = "hotmaps_task_2.7_load_profile_industry_iron_and_steel_generic.csv"
        case "ind:berg":
            load_profile_file_name = "hotmaps_task_2.7_load_profile_industry_non_metalic_minerals_generic.csv"
        case "ind:chem":
            load_profile_file_name = "hotmaps_task_2.7_load_profile_industry_chemicals_and_petrochemicals_generic.csv"
        case "tert":
            load_profile_file_name = "hotmaps_task_2.7_load_profile_tertiary_heating_generic.csv"
        case "wohn":
            load_profile_file_name = "hotmaps_task_2.7_load_profile_residential_heating_generic.csv"
        case _:
            raise Exception(f"Kein Standardlastprofil für den Gebäudetyp {building_type} gefunden. Verfügbare Gebäudetypen: wohn, tert, ind:papier, ind:leb, ind:stahl, ind:berg, ind:chem")
    load_profile = pd.read_csv(os.path.join(var_load_profiles.load_profile_dir, "generic", load_profile_file_name))
    return load_profile

def load_load_profile_shw(var_load_profiles, building_type: str) -> pd.DataFrame:
    """Loads the SHW load profile for the given building type.

    :param building_type: Indicates the type of building for which the 
                          load profile should be loaded. See match statement.
    :type building_type: str
    :return: The load profile.
    :rtype: pd.DataFrame
    """    
    match building_type:
        case "ind:papier":
            raise Exception(f"No load profile for SHW for building type {building_type} available.")
        case "ind:leb":
            raise Exception(f"No load profile for SHW for building type {building_type} available.")
        case "ind:stahl":
            raise Exception(f"No load profile for SHW for building type {building_type} available.")
        case "ind:berg":
            raise Exception(f"No load profile for SHW for building type {building_type} available.")
        case "ind:chem":
            raise Exception(f"No load profile for SHW for building type {building_type} available.")
        case "tert":
            load_profile_file_name = "hotmaps_task_2.7_load_profile_tertiary_shw_generic.csv"
        case "wohn":
            load_profile_file_name = "hotmaps_task_2.7_load_profile_residential_shw_generic.csv"
    load_profile = pd.read_csv(os.path.join(var_load_profiles.load_profile_dir, "generic", load_profile_file_name))
    return load_profile

def create_hourly_timestamps(year: int) -> pd.DataFrame:
    """Creates a dataframe with hourly timestamps for the given year.

    :param year: Year for which timestamps are to be calculated
    :type year: int
    :return: The timestamps for the given year
    :rtype: pd.DataFrame
    """    
    start_date = pd.Timestamp(year=year, month=1, day=1, hour=0)
    end_date = pd.Timestamp(year=year, month=12, day=31, hour=23)
    hourly_timestamps = pd.date_range(start_date, end_date, freq="H")
    hourly_timestamps = pd.DataFrame(hourly_timestamps, columns=["timestamp"])
    return hourly_timestamps

def process_load_profile(building_type: str, yearly_load: float, var_load_profiles, heating: bool) -> pd.DataFrame:
    """Customizes a generic load profile for a consumer based on the
    consumer's yearly load and the building type. The generic load
    profile is customized to the consumer's NUTS region and the year and
    normalized to the yearly load. The output is not a full hourly profile,
    but a scaled, sparse version of the generic profile.

    :param building_type: Type of the building, e.g. "ind:pap". See note in
                          the topology Excel sheet.
    :type building_type: str
    :param yearly_load: Yearly load the profile is to be normalized to [kWh]
    :type yearly_load: float
    :param var_load_profiles: Class containing all variables for the load profile generation
    :type var_load_profiles: var_load_profiles object
    :param heating: Set to True if the load profile is for heating, False if it is for SHW
    :type heating: bool
    :return: Scaled sparse load profile for the consumer
    :rtype: pd.DataFrame
    """

    NUTS_region = var_load_profiles.NUTS_region
    year = var_load_profiles.year

    # LOAD HOTMAPS PROFILE ####################################################
    ###########################################################################
    if heating:
        load_profile_generic_heating = load_load_profile_heating(var_load_profiles, building_type)
    if not heating:
        load_profile_generic_shw = load_load_profile_shw(var_load_profiles, building_type)

    # LOAD WEATHER DATA #######################################################
    ###########################################################################
    if heating:    
        weather_data = pd.read_csv(var_load_profiles.weather_file)
        # Cut the last 6 characters from "time" ("+00:00")
        weather_data["time"] = weather_data["time"].str[:-6]
        weather_data["time"] = pd.to_datetime(weather_data["time"])
    
    # PROCESSING ##############################################################
    ###########################################################################

    # INDUSTRIAL ##############################################################
    # Industrial profiles do not take outside temp. into account.
    # Also, there is no SHW profile for industry.
    if "ind" in building_type:
        # Filter load profile for the given NUTS region
        load_profile_generic_heating = load_profile_generic_heating[load_profile_generic_heating["NUTS0_code"] == NUTS_region[:2]]
        # Create hourly time stamps for the given year
        hourly_profile = create_hourly_timestamps(year)
        
        # Iterate through the time stamps
        for index, row in hourly_profile.iterrows():
            # Get the month and hour of the time stamp
            month = int(row["timestamp"].month)
            hour = int(row["timestamp"].hour)
            # Load profile hour naming conventions are different from the ones used in pandas
            if hour == 0:
                hour = 24
            # Get the daytype of the time stamp (0: weekday, 1: saturday, 2: sunday)
            daytype = row["timestamp"].dayofweek
            daytype = 0 if daytype <= 4 else 1 if daytype == 5 else 2
            # Extract the load for each time stamp
            hourly_profile.loc[index, "power [kW]"] = load_profile_generic_heating[(load_profile_generic_heating["month"] == month) & \
                                                                                   (load_profile_generic_heating["hour"] == hour) & \
                                                                                   (load_profile_generic_heating["daytype"] == daytype)]["load"].values[0]
        
        # Calculate the scale factor
        scale_factor = get_scale_factor(hourly_profile, yearly_load)
        # Scale the generic load profile
        load_profile_generic_heating["load"] = load_profile_generic_heating["load"] * scale_factor
    
    # TERTIARY AND RESIDENTIAL ################################################
    # Tertiary and residential profiles take outside temp. into account for heating.
    # For SHW, they take season into account. Also, the daytype column is called
    # 'day_type' instead of 'daytype' in the generic profiles. This is unified here.
    if "wohn" in building_type or "tert" in building_type:
        # Filter load profile for the given NUTS region
        if heating:
            load_profile_generic_heating = load_profile_generic_heating[load_profile_generic_heating["NUTS2_code"] == NUTS_region]
        if not heating:
            load_profile_generic_shw = load_profile_generic_shw[load_profile_generic_shw["NUTS2_code"] == NUTS_region]
        # Create hourly time stamps for the given year
        hourly_profile = create_hourly_timestamps(year)
        
        # Iterate through the time stamps
        cnt_errormessage = 0
        for index, row in hourly_profile.iterrows():
            # Get the hour of the time stamp
            hour = int(row["timestamp"].hour)
            # Load profile hour naming conventions are different from the ones used in pandas
            if hour == 0:
                hour = 24
            # Get the daytype of the time stamp (0: weekday, 1: saturday, 2: sunday)
            daytype = row["timestamp"].dayofweek
            daytype = 0 if daytype <= 4 else 1 if daytype == 5 else 2
            # Get the season of the time stamp:
            # 0: summer (15.5. to 14.9.)
            # 1: winter (1.11. to 20.3.)
            # 2: transitional (21.3. to 14.5. and 15.9. to 31.10.)
            month = int(row["timestamp"].month)
            start_summer = pd.Timestamp(year=year, month=5, day=15)
            end_summer = pd.Timestamp(year=year, month=9, day=14)
            start_winter = pd.Timestamp(year=year, month=11, day=1)
            end_winter = pd.Timestamp(year=year, month=3, day=20)
            if start_summer <= row["timestamp"] <= end_summer:
                season = 0
            elif start_winter <= row["timestamp"] <= end_winter:
                season = 1
            else:
                season = 2    

            if heating:
                # Get the temperature at the time stamp from the weather file
                temperature = weather_data.loc[weather_data["time"] == row["timestamp"], "TTX"].values[0]
                temperature = round(temperature)
                
                # Extract the load for each time stamp
                if temperature > 17:
                    # Set load to 0 for temperatures above the hotmaps specified range
                    hourly_profile.loc[index, "power [kW]"] = 0
                elif temperature < -15:
                    # Set temperature to minimum for temperatures below hotmaps specified range
                    temperature = -15
                    hourly_profile.loc[index, "power [kW]"] = load_profile_generic_heating[(load_profile_generic_heating["hour"] == hour) & \
                                                                                       (load_profile_generic_heating["day_type"] == daytype) & \
                                                                                       (load_profile_generic_heating["temperature"] == temperature)]["load"].values[0]
                else:
                    hourly_profile.loc[index, "power [kW]"] = load_profile_generic_heating[(load_profile_generic_heating["hour"] == hour) & \
                                                                                       (load_profile_generic_heating["day_type"] == daytype) & \
                                                                                       (load_profile_generic_heating["temperature"] == temperature)]["load"].values[0]
            if not heating:
                # Extract the load for each time stamp
                hourly_profile.loc[index, "power [kW]"] = load_profile_generic_shw[(load_profile_generic_shw["hour"] == hour) & \
                                                                                   (load_profile_generic_shw["day_type"] == daytype) & \
                                                                                   (load_profile_generic_shw["season"] == season)]["load"].values[0]
            
        # Calculate the scale factor
        scale_factor = get_scale_factor(hourly_profile, yearly_load)
        # Scale the generic load profile
        if heating:
            load_profile_generic_heating["load"] = load_profile_generic_heating["load"] * scale_factor 
            # Rename "day_type" column to "daytype"
            load_profile_generic_heating.rename(columns={"day_type": "daytype"}, inplace=True)
        if not heating:
            load_profile_generic_shw["load"] = load_profile_generic_shw["load"] * scale_factor  
            # Rename "day_type" column to "daytype"
            load_profile_generic_shw.rename(columns={"day_type": "daytype"}, inplace=True)
            
    # RETURNS #################################################################
    ###########################################################################
    if heating:
        return load_profile_generic_heating
    if not heating:
        return load_profile_generic_shw
    
def est_yearly_load_heating(eek: str, area: float) -> float:
    """Estimates the yearly heating load for a building with the given
    Energieeffizienzklasse (EEK) and area.
    
    Values taken from:
    https://www.verbraucherzentrale.de/wissen/energie/energetische-sanierung/energieausweis-was-sagt-dieser-steckbrief-fuer-wohngebaeude-aus-24074 
    (10.10.2023)

    :param eek: Energieeffizienzklasse (A+ ... H)
    :type eek: str
    :param area: Area of the building in m2
    :type area: float
    :return: The yearly load estimate in kWh.
    :rtype: float
    """    

    match eek:
        case "A+":
            yearly_load = 20 * area
        case "A":
            yearly_load = (30 + 50) / 2 * area
        case "B":
            yearly_load = (50 + 75) / 2 * area
        case "C":
            yearly_load = (75 + 100) / 2 * area
        case "D":
            yearly_load = (100 + 130) / 2 * area
        case "E":
            yearly_load = (130 + 160) / 2 * area
        case "F":
            yearly_load = (160 + 200) / 2 * area
        case "G":
            yearly_load = (200 + 250) / 2 * area
        case "H":
            yearly_load = 300 * area
        case _:
            raise Exception(f"Unbekannte EEK: {eek}")
    return yearly_load

def calc_kWh_m3(hist_data: pd.DataFrame) -> tuple[float, float]:
    """Calculates the kWh/m3 value from the given historical data.

    :param hist_data: Historical data
    :type hist_data: pd.DataFrame
    :return: kWh/m3
    :rtype: float
    """    
    # Flag rows where "Vorlauftemperatur (°C)" has not changed
    hist_data['Flag Unverändert'] = (hist_data['Vorlauftemperatur (°C)'] == hist_data['Vorlauftemperatur (°C)'].shift(periods=equal_values_max_timesteps - 1)) & \
            (hist_data['Ruecklauftemperatur (°C)'] == hist_data['Ruecklauftemperatur (°C)'].shift(periods=equal_values_max_timesteps - 1))
    # Filter for valid timestamps
    hist_data_filtered = hist_data[hist_data["Flag Unverändert"] == False]
    # Exclude NaNs
    hist_data_filtered = hist_data_filtered.dropna(subset=["Vorlauftemperatur (°C)", "Ruecklauftemperatur (°C)", "ges. Waermemenge (kWh)"])
    # Calculate "Durchfluss (m³/h)"
    hist_data_filtered["Durchfluss (m³/h)"] = hist_data_filtered["Durchfluss (l/h)"] / 1000
    # Calculate kWh/m3 for each timestamp
    hist_data_filtered["kWh/m³ (hist.)"] = hist_data_filtered["akt. Leistung(kW)"] / hist_data_filtered["Durchfluss (m³/h)"]
    # Calculate average kWh/m3
    kWh_m3 = hist_data_filtered["kWh/m³ (hist.)"].mean()
    t_VL = hist_data_filtered["Vorlauftemperatur (°C)"].mean()
    t_RL = hist_data_filtered["Ruecklauftemperatur (°C)"].mean()
    delta_T = t_VL - t_RL

    return kWh_m3, delta_T

# BOOKMARK: MAIN PROGRAM ######################################################
###############################################################################
###############################################################################
###############################################################################

# DELETE LOG FILE #############################################################
###############################################################################
###############################################################################

# Delete the log file if it exists
log_file = os.path.join(var_load_profiles.load_profile_dir, "lp_info.txt")
if os.path.exists(log_file):
    os.remove(log_file)

# READ TOPOLOGY FILE ########################################################
###############################################################################
###############################################################################
dtypes_topology = {"Nr.": str, "aktiv": str, "X": float, "Y": float, "h [m]": float,
                    "Verteiler": str, "p_ref": str, "Einspeiser": str, "Abnehmer": str,
                    "Nennleistung [kW]": float, "Energiebedarf [kWh/a]": float, "Durchsatz [m³/a]": float, "Gebäude": str,
                    "CSV": str, "Netzplan ID": str, "Inbetriebnahme": str, "Gebäudetyp": str,
                    "Hist. Daten existieren": str, "Wohn-/Nutzfläche [m²]": float,
                    "Anzahl Personen": int, "EEK": str, "Bemerkung": str}
topology_all = pd.read_excel(var_load_profiles.topology_file, sheet_name="Knoten", dtype=dtypes_topology)

# Filter for rows with a consumer ID, delete unnamed rows
topology = topology_all.dropna(subset=["Abnehmer"])

# Add new column "deltaT_mean" to topology
topology["deltaT_mean"] = np.nan

# INITIALIZE DIAGNOSTICS LISTS ################################################
###############################################################################
###############################################################################
l_hist_data_available = []
num_cons_hist_data_insufficient = 0

# READ GENERAL_INFO FILE ######################################################
###############################################################################
###############################################################################
# Read general_info file
path = os.path.join(var_load_profiles.cons_dir, "general_info.csv")
general_info = pd.read_csv(path, delimiter=";", dtype={"storages": str})
# Each season (winter, transitional, summer) spans four months. Therefore, we
# can just take the average of the 3 to get an average daily load for SHW
# and heating.

# Iterate through IDs in topology file
print_blue("Lese Warmwasser-/Heizungsverbrauchsinformationen.")
heating_perc_dict = {}
for cons_id in topology["Abnehmer"]:
    # Find the matching entry in general_info ("storages" column)
        current_line = general_info[general_info["storages"] == cons_id]
        # If there is no entry, set a flag and continue
        if len(current_line) == 0:
            print_yellow(f"Keine Verbrauchsinformationen für Abnehmer {cons_id} gefunden. Das Verhältnis Warmwasser:Heizungsverbrauch wird auf den Durchschnitt der anderen Abnehmer gesetzt.")
            heating_perc = "TBD"
        # If there is an entry, extract the heating percentage
        else:
            avg_daily_consumption_heating = np.mean([current_line["avg. daily consumption heating winter"].values[0],
                                                     current_line["avg. daily consumption heating transition"].values[0],
                                                     current_line["avg. daily consumption heating summer"].values[0]])
            avg_daily_consumption_shw = np.mean([current_line["avg. daily consumption storage winter"].values[0],
                                                 current_line["avg. daily consumption storage transition"].values[0],
                                                 current_line["avg. daily consumption storage summer"].values[0]])
            heating_perc = avg_daily_consumption_heating / (avg_daily_consumption_heating + avg_daily_consumption_shw)
            # If the heating percentage is NaN, set a flag and continue
            if np.isnan(heating_perc):
                print_yellow(f"Keine Verbrauchsinformationen für Abnehmer {cons_id} \
                             gefunden. Das Verhältnis Warmwasser:Heizungsverbrauch \
                             wird auf den Durchschnitt der anderen Abnehmer gesetzt.")
                heating_perc = "TBD"
                continue
        # Add cons_id and heating_perc to a dict
        heating_perc_dict[cons_id] = heating_perc

# Fill all consumers where no percentage could be calculated with the average
avg_heating_perc = np.mean([value for value in heating_perc_dict.values() if value != "TBD"])
for cons_id in heating_perc_dict.keys():
    if heating_perc_dict[cons_id] == "TBD":
        heating_perc_dict[cons_id] = avg_heating_perc

# Make another dict for shw percentages
shw_perc_dict = {}
for cons_id in heating_perc_dict.keys():
    shw_perc_dict[cons_id] = 1 - heating_perc_dict[cons_id]

# READ CONSUMER FILES #########################################################
###############################################################################
###############################################################################
# Iterate through IDs in topology file
print_blue("Lese Abnehmerdateien.")
cntr_processed = 0
cntr_total = len(topology)
for index, row in topology.iterrows():
    #  INITIALIZATION #########################################################
    ###########################################################################
    # Initialize load profiles
    if 'load_profile_heating' in locals():
        del load_profile_heating
    if 'load_profile_shw' in locals():
        del load_profile_shw
    # Skip inactive consumers
    if "nein" in row["aktiv"]:
        print(" ")
        print("-----------------------------------------------------------------------------------------------------------")
        print(f"ABNEHMER {row['Abnehmer']} ist inaktiv. Fahre fort ohne Lastprofilberechnung.")
        print("-----------------------------------------------------------------------------------------------------------")
        print(" ")
        cntr_total -= 1
        continue

    cntr_processed += 1

    cons_id = row["Abnehmer"]
    print("/```````````````````````````````\__________________________________________________________________________")
    print(f"     ABNEHMER {cons_id} --- {cntr_processed:02d}/{cntr_total}")
    try:
        building_type = row["Gebäudetyp"] # residential, tertiary, industrial, ...
        print(f"     Gebäudetyp: {building_type}")
        print("")
    except:
        raise Exception(f"Abnehmer {cons_id}: Gebäudetyp nicht angegeben. Bitte in Netztopologie eintragen.")

    # READ HISTORICAL DATA ####################################################
    ###########################################################################
    file_path = os.path.join(var_load_profiles.cons_dir, 'Regler_' + cons_id + '_prepared.csv')
    # Check if there is a consumer file for the current ID. If not, set a flag.
    try:
        hist_data = pd.read_csv(file_path)
        row["Hist. Daten existieren"] = True
    except:
        print_red(f"Die Datei {file_path} konnte nicht gefunden werden. Fahre fort ohne hist. Daten.")
        print("---")
        row["Hist. Daten existieren"] = False

    # CHECK DATA QUALITY ######################################################
    ###########################################################################
    # If more than 10% of last winter's data are missing, set a flag.
    if row["Hist. Daten existieren"]:
        # Rename first column to "Zeit"
        hist_data = hist_data.rename(columns={hist_data.columns[0]: "Zeit"})
        # Convert first column to datetime
        hist_data["Zeit"] = pd.to_datetime(hist_data["Zeit"])

        # Filter hist_data for the winter periods of the given year
        start_first_winter = pd.to_datetime(f"{var_load_profiles.year}-01-01 00:00:00")
        end_first_winter = pd.to_datetime(f"{var_load_profiles.year}-03-31 23:45:00")
        start_second_winter = pd.to_datetime(f"{var_load_profiles.year}-10-01 00:00:00")
        end_second_winter = pd.to_datetime(f"{var_load_profiles.year}-12-31 23:45:00")

        hist_data_winter = hist_data[((hist_data["Zeit"] >= start_first_winter) & (hist_data["Zeit"] <= end_first_winter)) | \
                                        ((hist_data["Zeit"] >= start_second_winter) & (hist_data["Zeit"] <= end_second_winter))]

        # Check if number of data dropouts in the last winter exceeds threshold of one day's worth of dropouts
        equal_values_max_timesteps = int(var_gaps.equal_values_max_min / 15)

        # Flag rows where "Vorlauftemperatur (°C)" has not changed
        hist_data_winter['Flag Unverändert'] = (hist_data_winter['Vorlauftemperatur (°C)'] == hist_data_winter['Vorlauftemperatur (°C)'].shift(periods=equal_values_max_timesteps - 1)) & \
            (hist_data_winter['Ruecklauftemperatur (°C)'] == hist_data_winter['Ruecklauftemperatur (°C)'].shift(periods=equal_values_max_timesteps - 1))

        # Count the number of dropouts
        num_dropouts = hist_data_winter["Flag Unverändert"].sum()
        perc_dropouts = num_dropouts / len(hist_data_winter) * 100

        if len(hist_data_winter) < 10000:
            # If there is no data in the winter, set a flag and print a warning
            row["Hist. Daten existieren"] = False
            print_yellow(f"Keine Daten im Winter gefunden - Berechnung d. Jahresgesamtverbrauchs mit Daten aus dem Netztopologie-File.")
            print_yellow("Ein Standardlastprofil für Heizung und SHW wird erstellt. Der kWH/m³-Wert wird nicht berechnet.")
            print("---")
            num_cons_hist_data_insufficient += 1
        
        if perc_dropouts > var_load_profiles.threshold_data_dropouts:
            # If so, set a flag and print a warning
            row["Hist. Daten existieren"] = False
            print_yellow(f"> {var_load_profiles.threshold_data_dropouts} % Datenlücken im Winter - Berechnung d. Jahresgesamtverbrauchs mit Daten aus dem Netztopologie-File.")
            print_yellow("Ein Standardlastprofil für SHW wird erstellt. Der kWH/m³-Wert wird nicht berechnet.")
            print("---")
            num_cons_hist_data_insufficient += 1
    else:
        pass

    # Fill NaN values in hist_data["Aussentemperatur (°C)"] with the last valid value
    hist_data["Aussentemperatur (°C)"].ffill(inplace=True)

    # CALCULATE LOAD PROFILES #################################################
    ###########################################################################

    # HISTORICAL DATA AVAILABLE ###############################################
    # Only calculate heating from gen. load profiles; SHW will be copied from hist. data
    if row["Hist. Daten existieren"]:
        """
        yearly_load = hist_data[hist_data["Zeit"] == pd.to_datetime(f"{year+1}-01-01 00:05:00")]["ges. Waermemenge (kWh)"].values[0] - \
        hist_data[hist_data["Zeit"] == pd.to_datetime(f"{year}-01-01 00:05:00")]["ges. Waermemenge (kWh)"].values[0]
        NOTE: The above, which would be the correct way to calculate the yearly load, does not work because the "ges. Waermemenge (kWh)" 
        NOTE: column resets for some consumers in the testing data. Instead, the yearly load is calculated from the quarter-hourly "Leistung (kW)" values.
        """
        print("Historische Daten vorhanden.")
        print(f"Berechnung d. Jahresgesamtverbrauchs mit Daten aus {var_load_profiles.year}.")
        print("---")

        # Filter for current year
        hist_data = hist_data[hist_data["Zeit"].dt.year == var_load_profiles.year]
        # Calculate yearly load in kWh from hist. data
        yearly_load = hist_data[hist_data["Zeit"].dt.year == var_load_profiles.year]["Heizung (kW)"].sum() * 0.25 
        # Create a customized load profile for the consumer
        load_profile_heating = process_load_profile(building_type, yearly_load, var_load_profiles, heating=True)

        # Calculate kWh/m3 from hist. data
        print_green("Berechnung kWh/m³ aus historischen Daten.")
        row["kWh/m³ (hist.)"], row["deltaT_mean"] = calc_kWh_m3(hist_data)

    # BOOKMARK: HISTORICAL DATA UNAVAILABLE OR INSUFFICIENT #############################
    # Calculate heating and shw from generic load profiles. For industry and 
    # tertiary, only heating is calculated.
    else:
        # HEATING #
        # Check if "Energiebedarf [kWh/a]" and "Durchsatz [m³/a]" are available
        if not pd.isna(row["Energiebedarf [kWh/a]"]) and not pd.isna(row["Durchsatz [m³/a]"]):
            row["kWh/m³ (hist.)"] = row["Energiebedarf [kWh/a]"] / row["Durchsatz [m³/a]"]
            print_green("Berechnung kWh/m³ aus Energiebedarf und Durchsatz.")
            # If consumer is residential, split the yearly load into heating and shw
            if "wohn" in building_type:
                yearly_load = row["Energiebedarf [kWh/a]"] * heating_perc_dict[cons_id]
            else: 
                yearly_load = row["Energiebedarf [kWh/a]"]
        # If not, calculate yearly load from EEK and area    
        else:
            try:
                eek = row["EEK"]
            except:
                raise Exception(f"Abnehmer {cons_id}: EEK nicht angegeben. Bitte in Netztopologie eintragen.")
            try:
                area = row["Wohn-/Nutzfläche [m²]"]
            except:
                raise Exception(f"Abnehmer {cons_id}: Wohn-/Nutzfläche nicht angegeben. Bitte in Netztopologie eintragen.")
            # Check if the area is a valid number (positive, not NaN, not 0)
            if area <= 0 or pd.isna(area):
                raise Exception(f"Abnehmer {cons_id}: Ungültige Wohn-/Nutzfläche. Bitte in Netztopologie eintragen.")
            # Check if area is a number
            if not isinstance(area, (int, float, np.float64)):
                raise Exception(f"Abnehmer {cons_id}: Ungültige Wohn-/Nutzfläche. Bitte in Netztopologie eintragen.")
            # Calculate yearly load estimate from EEK and area
            yearly_load = est_yearly_load_heating(eek, area)
        # Create a customized load profile for the consumer
        load_profile_heating = process_load_profile(building_type, yearly_load, var_load_profiles, heating=True)

        
        # SHW #

        # Calculate only for residential buildings
        if "wohn" in building_type:
            # Check if "Energiebedarf [kWh/a]" and "Durchsatz [m³/a]" are available
            if not pd.isna(row["Energiebedarf [kWh/a]"]) and not pd.isna(row["Durchsatz [m³/a]"]):
                yearly_load = row["Energiebedarf [kWh/a]"] * shw_perc_dict[cons_id]
            # If not, calculate yearly load from number of residents
            else:
                try:
                    num_persons = row["Anz. Personen"]
                except:
                    raise Exception(f"Abnehmer {cons_id}: Anzahl Personen nicht angegeben. Bitte in Netztopologie eintragen.")
                # Check if the number of persons is a valid number (positive, not NaN, not 0)
                if num_persons <= 0 or pd.isna(num_persons):
                    raise Exception(f"Abnehmer {cons_id}: Ungültige Anzahl Personen (muss eine positive Zahl sein). Bitte in Netztopologie eintragen.")
                # Check if number of persons is a number
                if not isinstance(num_persons, (int, float, np.float64)):
                    raise Exception(f"Abnehmer {cons_id}: Ungültige Anzahl Personen. Bitte in Netztopologie eintragen.")
                # The energy demand for SHW per person is approx. 1000 kWh/a as per
                # https://www.umweltberatung.at/download/?id=warmes-wasser-3073-umweltberatung.pdf
                yearly_load = num_persons * 1000
            # Create a customized load profile for the consumer
            load_profile_shw = process_load_profile(building_type, yearly_load, var_load_profiles, heating=False)

    # WRITE TOPOLOGY INFO TO DATAFRAME ########################################
    ###########################################################################
    topology.loc[index] = row

    # WRITE LOAD PROFILES TO FILE #############################################
    ###########################################################################
    try:
        load_profile_heating.to_csv(os.path.join(var_load_profiles.load_profile_dir, 'individual', 'Regler_' + cons_id + '_heating.csv'), index=False)
        print_green("Lastprofil für Heizung gespeichert.")
    except:
        print("Kein Lastprofil für Heizung gespeichert.")

    try:
        load_profile_shw.to_csv(os.path.join(var_load_profiles.load_profile_dir, 'individual', 'Regler_' + cons_id + '_shw.csv'), index=False)
        print_green("Lastprofil für Warmwasser gespeichert.")
    except:
        print("Kein Lastprofil für Warmwasser gespeichert.")
    
    print(" ")

# WRITE TOPOLOGY FILE #########################################################
###############################################################################
# Save historical data availability to file
print_blue(f"Lastprofile wurden berechnet. Für {num_cons_hist_data_insufficient} Abnehmer waren die historischen Daten nicht ausreichend.")
print_blue("Schreibe Ergebnisse in Netztopologie.")
topo_wb = openpyxl.load_workbook(var_load_profiles.topology_file)
topo_ws = topo_wb["Knoten"]
# Write "Hist. Daten existieren" and "kWh/m³ (hist.)" to columns R and V of sheet
for index, row in topology.iterrows():
    topo_ws[f"R{index + 2}"] = row["Hist. Daten existieren"]
    topo_ws[f"V{index + 2}"] = row["kWh/m³ (hist.)"]

topo_wb.save(var_load_profiles.topology_file)

# WRITE YEAR INFO TO FILE #####################################################
###############################################################################
# Write the year and the name of input excel to a file
with open(os.path.join(var_load_profiles.load_profile_dir, "lp_info.txt"), "w") as file:
    file.write(str(var_load_profiles.year))
    file.write("\n")
    file.write(var_load_profiles.topology_file)

# SHOW DIAGNOSTICS ############################################################
###############################################################################
if var_load_profiles.show_regression_plot:
    plt.figure()
    sns.regplot(x="deltaT_mean", y="kWh/m³ (hist.)", data=topology, scatter_kws={"color": "blue"}, line_kws={"color": "red"})
    plt.xlabel("Mittlere Temperaturdifferenz [°C]")
    plt.ylabel("kWh/m³")
    plt.show()


