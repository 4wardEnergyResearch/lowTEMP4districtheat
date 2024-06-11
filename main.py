# -*- coding: utf-8 -*-
"""
Author: 4wardEnergy Research GmbH
Date: 2024-05-14
Version: 1.0

This script is designed to execute a comprehensive simulation of a hydraulic network.
It integrates various modules to perform data preparation, input validation, and the
execution of the main simulation program. The script also includes functionalities for
ML model calculation and input file handling.

Functions:
- print_red(text): Prints the provided text in red color.
- button_action(): Executes the main program with the provided input parameters.
- button_select(): Opens a file selection dialog and updates the input file information.
- button_data_prep(): Initiates the data preparation process.
- button_calc_ml_models(): Opens a window for ML model calculation.
- ml_no_file_selected(message_window): Handles the case when no file is selected for ML model calculation.
- change_calc_all_label(calc_all, calc_all_label): Updates the label text based on the state of the 'calculate all' checkbox.
- start_ml_model_calculation(overwrite, calc_all, ml_window): Executes the ML model calculation.
- check_ml_model_availability(file_input): Checks the availability and last modified dates of necessary ML models.
- on_year_slp_change(event=None): Compares the input year with the current SLP year and updates the label color accordingly.
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

import os, sys
project_dir = os.path.dirname(os.path.abspath(__file__))
# Add the project directory to the system path
sys.path.append(project_dir)
# Set working directory to project directory
os.chdir(project_dir)

import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from tkinter import *
from tkinter.filedialog import askopenfilename
from datetime import timedelta, datetime
from openpyxl.styles import NamedStyle
from simulation import main_program
import matplotlib
import math
from options import *
import termcolor
import pandas as pd
import importlib




def print_red(text):
    print(termcolor.colored(text, 'red'))


# DELETE ALL VARIABLES
#sys.modules[__name__].__dict__.clear()
#%reset -f

matplotlib.use('TkAgg')



###############################################################################
# INTIALIZE VARIABLES #########################################################
###############################################################################

# Time stamps for hydraulic calculation
var_sim.time_stamp = []

# Maximum number of unchanged temperature values to detect an error [-]
var_gaps.nbr_equal_values_max = math.ceil(var_gaps.equal_values_max_min / var_sim.delta_time_hyd)

# Initialization of various variables
plots.figure4, plots.bar4, plots.cb, plots.fig, plots.fig2 = ("" for i in range (5))

# UNUSED VARIABLES (LEGACY) ###################################################
class var_unused:
    time_excel_sum, delta_Q_dot_percent, sheet_delta_Q_dot = ([] for i in range(3))


def button_action():
    """Executes the main program.
    """
    # Change Label of button
    change_button.config(text="In progress...")
    # Update input window
    input_window.update()

    # Check if end time is after start time
    if datetime.strptime(input_box_end.get(), "%d.%m.%Y %H:%M") <= datetime.strptime(input_box_start.get(), "%d.%m.%Y %H:%M"):
        print_red("The end time must be after the start time.")
        # Open error window
        error_window = Toplevel()
        error_window.title("Fehler")
        error_window.geometry("300x100")
        error_label = Label(error_window, text="The end time must be after the start time.")
        error_label.pack(pady=10)
        ok_button = Button(error_window, text="OK", command=error_window.destroy)
        ok_button.pack(pady=10)
        # Change Label of button back to "Simulation starten"
        change_button.config(text="Start Simulation")
        # Update input window
        input_window.update()
        return

    # SET START TIME OF SIMULATION ###########################################
    # Set start time to last available time step before input time
    time_sim_start_input = datetime.strptime(input_box_start.get(), "%d.%m.%Y %H:%M")
    var_sim.time_sim_start = datetime.strptime("01.01.2020 00:05", "%d.%m.%Y %H:%M")

    if var_sim.time_sim_start < time_sim_start_input:
    # Input time after 01.01.2020: Search forward
        while var_sim.time_sim_start <= time_sim_start_input:
            var_sim.time_sim_start = var_sim.time_sim_start + timedelta(minutes=var_sim.delta_time_hyd)
        var_sim.time_sim_start = var_sim.time_sim_start - timedelta(minutes=var_sim.delta_time_hyd)
    else:
    # Input time before 01.01.2020: Search backward
        while var_sim.time_sim_start > time_sim_start_input:
            var_sim.time_sim_start = var_sim.time_sim_start - timedelta(minutes=var_sim.delta_time_hyd)
    var_sim.time_stamp.append(var_sim.time_sim_start)

    # SET END TIME OF SIMULATION #############################################
    # Set end time to first available time step after input time
    time_sim_end_input = datetime.strptime(input_box_end.get(), "%d.%m.%Y %H:%M")
    var_sim.time_sim_end = datetime.strptime("01.01.2020 00:05", "%d.%m.%Y %H:%M")

    # If real-time checkbox is ticked, ignore end time input and set end time to now
    
    if var2.get() == 1:
        var_sim.real_time = True
        time_sim_end_input = datetime.now()
    else:
        var_sim.real_time = False

    if var_sim.time_sim_end < time_sim_end_input:
        while var_sim.time_sim_end < time_sim_end_input:
            var_sim.time_sim_end = var_sim.time_sim_end + timedelta(minutes=var_sim.delta_time_hyd)
    else:
        while var_sim.time_sim_end >= time_sim_end_input:
            var_sim.time_sim_end = var_sim.time_sim_end - timedelta(minutes=var_sim.delta_time_hyd)
        var_sim.time_sim_end = var_sim.time_sim_end + timedelta(minutes=var_sim.delta_time_hyd)

    # SET TIME STAMPS ########################################################
    # Set time stamps for hydraulic simulation every 15 minutes
    while var_sim.time_stamp[-1] < var_sim.time_sim_end:
        var_sim.time_stamp.append(var_sim.time_stamp[-1] + timedelta(minutes=var_sim.delta_time_hyd))

    # TOTAL NUMBER OF TIME STEPS IN SIMULATION
    var_sim.time_steps = len(var_sim.time_stamp)

    if var1.get() == 1:  # "Show Graphics: yes"
        plots.show_plot = "yes"
        # plots.ax4 = plt.axes()
    else:
        plots.show_plot = "no"
        # plots.ax4 = ""

    try:
        excel_save_time = int(excel_interval_entry.get()) 
    except:
        excel_save_time = 0

    if excel_save_time == 0 and var_sim.real_time == False:
        print("No caching time set. Caching disabled.")

    if var_sim.real_time == True:
        excel_save_time = 0.25
        print("Real-time simulation activated. Caching at every time step.")

    # EXECUTE MAIN_PROGRAM.PY
    main_program.main_program(var_phy, var_H2O, var_misc, var_sim, plots, \
                              file_input, excel_save_time, var_unused, var_gaps)

    print("End of Simulation.")
    sys.exit()

    return (node, line)

def button_select():
    """Opens a file selection window and returns the selected file.
    Also writes the file name to the options file."""
    global file_input
    file_input = askopenfilename()
    parted = file_input.split('/')
    input_file_name.config(state='normal')
    input_file_name.delete(0, END)
    input_file_name.insert(0, parted[-1])
    input_file_name.config(state='disabled')

    # Check if input_file_name is a valid .xlsx file
    if not file_input.endswith('.xlsx'):
        print("Invalid file. Only xlsx files are supported.")
        sys.exit()

    # Check if all necessary ml models are available
    ml_models_last_changed = check_ml_model_availability(file_input)
    total_needed_ml_models = len(ml_models_last_changed)
    if None in ml_models_last_changed:
        # Count how many ml models are missing
        missing_ml_models = ml_models_last_changed.count(None)
        ml_status_label.config(text=f"{missing_ml_models}/{total_needed_ml_models} ML-Models n/a.", fg="red")
        # Lock the "Bestätigen" button
        change_button.config(state='disabled')
    else:
        ml_status_label.config(text=f"Alle {total_needed_ml_models} ML-Models available.", fg="green")
        change_button.config(state='normal')
    
    # Find latest timestamp of ml models
    ml_models_last_changed = [x for x in ml_models_last_changed if x is not None]
    if len(ml_models_last_changed) > 0:
        latest_ml_model = max(ml_models_last_changed)
        ml_date_range_label.config(text=f"{latest_ml_model}")
    else:
        ml_date_range_label.config(text="No ML-Models available.")

    # Check if the file name corresponds with the current SLP file
    if os.path.basename(file_input) == slp_current_file:
        slp_current_file_label.config(fg="green")
    else:
        slp_current_file_label.config(fg="red")
    
    # Check if the year corresponds with the current SLP year
    if input_year_slp.get() == slp_current_year:
        slp_current_year_label.config(fg="green")
    else:
        slp_current_year_label.config(fg="red")


    # Replace the topology file in options.py with the selected file
    with open("options.py", "r") as f:
        lines = f.readlines()
    with open("options.py", "w") as f:
        for line in lines:
            if "topology_file =" in line:
                line = "    topology_file = '" + file_input + "'\n"
            f.write(line)

    return file_input

def button_data_prep():
    """Executes wrapper_data_preparation.py."""
    # Open info window
    info_window = Toplevel()
    info_window.title("Datenaufbereitung starten")
    info_window.geometry("340x150")
    # Message: "Datenaufbereitung starten?"
    info_label = Label(info_window, text="Data preparation may take a few minutes \n and is only necessary if the SLP year, \n the csvs, or the input file have been changed. \n Should data preparation be initiated?")
    info_label.pack(side="top", fill="x",pady=10,padx=10)
    # Button: "Ja"
    def start_data_prep():
        info_window.destroy()
        # Check if a file has been selected
        try:
            file_input
        except NameError:
            print_red("No file selected. Please select a file.")
            return

        # Read entry for year
        try:
            input_year_slp_int = int(input_year_slp.get())
            if input_year_slp_int < 2000 or input_year_slp_int > 2100:
                raise ValueError
        except ValueError:
            print_red("Invalid year. Please enter a year between 2000 and 2100.")
            return

        # Replace the file name in options.py with the selected file

        # Replace the year in options.py with the selected year
        with open("options.py", "r") as f:
            lines = f.readlines()
        with open("options.py", "w") as f:
            for line in lines:
                if "year: int = " in line:
                    line = "    year: int = " + str(input_year_slp_int) + " \n"
                f.write(line)

        # Change Label of button to "In Bearbeitung..."
        data_prep_button.config(text="In progress...")
        # Update input window
        input_window.update()

        # Execute data preparation
        import data_prep.wrapper_data_preparation as wrapper_data_preparation

        # Read info from "lp_info.txt"
        # Check if log file exists
        if not os.path.exists(os.path.join(var_load_profiles.load_profile_dir, "lp_info.txt")):
            slp_current_year = "No SLP available"
            slp_current_file = "No SLP available"
        else:
            with open(os.path.join(var_load_profiles.load_profile_dir, "lp_info.txt"), "r") as f:
                lines = f.readlines()
            slp_current_year = lines[0].strip()
            slp_current_file = lines[1].strip()
            slp_current_file = os.path.basename(slp_current_file)

        # Update labels
        slp_current_year_label.config(text=slp_current_year)
        slp_current_file_label.config(text=slp_current_file)

        # Check if the file name corresponds with the current SLP file
        if os.path.basename(file_input) == slp_current_file:
            slp_current_file_label.config(fg="green")
        else:
            slp_current_file_label.config(fg="red")
        
        # Check if the year corresponds with the current SLP year
        if input_year_slp.get() == slp_current_year:
            slp_current_year_label.config(fg="green")
        else:
            slp_current_year_label.config(fg="red")

        # Change Label of button back to "Datenaufbereitung"
        data_prep_button.config(text="Data preparation")
        # Update input window
        input_window.update()

    info_button_yes = Button(info_window, text="Yes", command=start_data_prep)
    info_button_yes.pack(side="left", padx=50, pady=10)
    # Button: "Nein"
    info_button_no = Button(info_window, text="No", command=info_window.destroy)
    info_button_no.pack(side="right", padx=50, pady=10)
    # Set focus to "Ja" button
    info_button_yes.focus()

    # Wait for user input
    info_window.mainloop()

    return

def button_calc_ml_models():
    """Opens a window to calculate ml models."""
    # New info window
    ml_window = Toplevel()
    ml_window.title("Calculate ML-Models")
    ml_window.geometry("450x280")

    # Check if a file has been selected
    try:
        file_input
    except NameError:
        # Display a message window
        message_window = Toplevel()
        # Bring window to front
        message_window.attributes("-topmost", True)
        message_window.title("No file selected")
        message_window.geometry("300x100")
        message_label = Label(message_window, text="No xlsx file selected. Please select a file.")
        message_label.pack(pady=10)
        ok_button = Button(message_window, text="OK", command=lambda:ml_no_file_selected(message_window))
        ok_button.pack(pady=10)
        return

    # Create info text
    info_label = Label(ml_window, text="The calculation of ML models takes about 30 min ~ 1 h per consumer. \n The hard drive space requirement is about 1 GB per consumer.", fg="red")
    calc_all_label = Label(ml_window, text="ML models are only calculated for those consumers \n for whom AI gap filling is intended in the input Excel.")
    question_label = Label(ml_window, text="Should the ML models be calculated now and the automatically \n chosen gapfilling methods updated?")

    # Create the checkboxes
    overwrite = IntVar()
    calc_all = IntVar()
    checkbox1 = Checkbutton(ml_window, text="Overwrite existing Models", variable=overwrite)
    checkbox2 = Checkbutton(ml_window, text="Calculate models for all consumers", variable=calc_all, command=lambda: change_calc_all_label(calc_all, calc_all_label))


    # Create the buttons
    calc_button = Button(ml_window, text="Calculate", command=lambda:start_ml_model_calculation(overwrite, calc_all, ml_window))
    cancel_button = Button(ml_window, text="Cancel", command=ml_window.destroy)

    # Position the checkboxes and buttons
    info_label.pack(pady=10)
    checkbox1.pack()
    checkbox2.pack()
    calc_all_label.pack(pady=10)
    question_label.pack(pady=10)
    calc_button.pack() 
    cancel_button.pack()

def ml_no_file_selected(message_window):
    """ Opens a file selection window and closes the "No file selected" message window."""
    # Close the window
    message_window.destroy()
    # Run button_select() to open a file selection window
    button_select()
    return

def change_calc_all_label(calc_all, calc_all_label):
    """ Changes the text of the calc_all_label depending on the state of the calc_all checkbox."""
    if calc_all.get() == 1:
        calc_all_label.config(text="ML models are calculated for all consumers.")
    else:
        calc_all_label.config(text="ML models are only calculated for those consumers \n for whom AI gap filling is intended in the input Excel.")


def start_ml_model_calculation(overwrite, calc_all, ml_window):
    """ Executes the ml model calculation and closes the ml_window."""
    overwrite_var = overwrite.get()
    calc_all_var = calc_all.get()

    if overwrite_var == 1:
        # Set "overwrite_ml_models" to True in options.py
        with open("options.py", "r") as f:
            lines = f.readlines()
        with open("options.py", "w") as f:
            for line in lines:
                if "overwrite_ml_models = " in line:
                    line = "    overwrite_ml_models = True \n"
                f.write(line)
    else:
        # Set "overwrite_ml_models" to False in options.py
        with open("options.py", "r") as f:
            lines = f.readlines()
        with open("options.py", "w") as f:
            for line in lines:
                if "overwrite_ml_models = " in line:
                    line = "    overwrite_ml_models = False \n"
                f.write(line)

    if calc_all_var == 1:
        # Set "force_ml_models" to True in options.py
        with open("options.py", "r") as f:
            lines = f.readlines()
        with open("options.py", "w") as f:
            for line in lines:
                if "force_ml_models = " in line:
                    line = "    force_ml_models = True \n"
                f.write(line)
    else:
        # Set "force_ml_models" to False in options.py
        with open("options.py", "r") as f:
            lines = f.readlines()
        with open("options.py", "w") as f:
            for line in lines:
                if "force_ml_models = " in line:
                    line = "    force_ml_models = False \n"
                f.write(line)
    
    # Import ml models module (This must be done here so the options are updated before the module is imported)
    from data_prep import data_prep_ml_models
    # Execute data_prep_ml_models.py main routine   
    data_prep_ml_models.main_routine()

    # Check if all necessary ml models are available
    ml_models_last_changed = check_ml_model_availability(file_input)
    total_needed_ml_models = len(ml_models_last_changed)
    if None in ml_models_last_changed:
        # Count how many ml models are missing
        missing_ml_models = ml_models_last_changed.count(None)
        ml_status_label.config(text=f"{missing_ml_models}/{total_needed_ml_models} ML models n/a.", fg="red")
        # Lock the "Bestätigen" button
        change_button.config(state='disabled')
    else:
        ml_status_label.config(text=f"Alle {total_needed_ml_models} ML models available.", fg="green")
        change_button.config(state='normal')

    # Display a message window
    message_window = Toplevel()
    message_window.title("ML models calculated")
    message_window.geometry("300x100")
    message_label = Label(message_window, text="ML models were calculated.")
    message_label.pack(pady=10)
    ok_button = Button(message_window, text="OK", command=message_window.destroy)
    ok_button.pack(pady=10)

    # Close the window
    ml_window.destroy()
    

def check_ml_model_availability(file_input):
    """ 
    Checks if all necessary ml models are available and returns their last changed dates.
    
    :param file_input: The file path of the input file.
    :type file_input: str
    
    :return: A list of last changed dates for each ml model. None if no ml model is available.
    :rtype: list
    """


    # Load the topology file
    topology_file = pd.read_excel(file_input, sheet_name="Knoten", dtype={"Abnehmer": str, "KI": str, "Lückenfüllung": str})

    # Filter for rows where "Abnehmer" is not none
    cons_list = topology_file[topology_file["Abnehmer"].notna()]

    # Filter for consumers where "aktiv" is "ja"
    cons_list = cons_list[cons_list["aktiv"] == "ja"]

    # Initialize lists
    ml_model_last_changed_ls = []
    
    # Iterate through rows of cons_list
    for index, row in cons_list.iterrows():
        lueckenfuellung = row["Lückenfüllung"]
        override = row["Override"]
        cons_id = row["Abnehmer"]
        # Check if override is NaN

        if not pd.isna(override):
            lueckenfuellung = override

        if lueckenfuellung != "KI":
            continue
        
        # Check if ml model is available
        path = os.path.join(var_ml_models.ml_model_dir, cons_id)
        # Check if the directory exists
        if not os.path.exists(path):
            ml_model_last_changed = None
        else:
            # List txt files beginning with "last_changed"
            txt_files = [f for f in os.listdir(path) if f.startswith("last_changed")]
            # Check if there are any txt files
            if len(txt_files) == 0:
                ml_model_last_changed = None
            else:
                ml_model_last_changed = txt_files[0]
                # Extract last changed date
                ml_model_last_changed = ml_model_last_changed[13:-4]
                # Convert to timestamp
                ml_model_last_changed = datetime.strptime(ml_model_last_changed, "%Y-%m-%d_%H-%M-%S")
        
        # Append to list
        ml_model_last_changed_ls.append(ml_model_last_changed)

    return ml_model_last_changed_ls

def on_year_slp_change(event=None):
    """ Compares the input year with the current SLP year and changes the label color accordingly."""
    if input_year_slp.get() == slp_current_year:
        slp_current_year_label.config(fg="green")
    else:
        slp_current_year_label.config(fg="red")

###############################################################################
###############################################################################
# MAIN PROGRAM ###############################################################
###############################################################################
###############################################################################


if __name__ == "__main__":
    # Read info from "lp_info.txt"
    # Check if log file exists
    if not os.path.exists(os.path.join(var_load_profiles.load_profile_dir, "lp_info.txt")):
        slp_current_year = "No SLPs available"
        slp_current_file = "No SLPs available"
    else:
        with open(os.path.join(var_load_profiles.load_profile_dir, "lp_info.txt"), "r") as f:
            lines = f.readlines()
        slp_current_year = lines[0].strip()
        slp_current_file = lines[1].strip()
        slp_current_file = os.path.basename(slp_current_file)

    # Create a window
    input_window = Tk()
    # Create the window title
    input_window.title("Network simulation")

    # CREATE LABELS AND BUTTONS
    change_button = Button(input_window, text="Start Simulation", command=button_action)
    select_button = Button(input_window, text="Select", command=button_select)
    data_prep_button = Button(input_window, text="Data preparation", command=button_data_prep)
    input_file_name = Entry(input_window, bd=5, width=15, state='disabled')
    input_year_slp = Entry(input_window, bd=5, width=15)
    input_year_slp.insert(0, "2022")
    input_box_start = Entry(input_window, bd=5, width=15)
    input_box_end = Entry(input_window, bd=5, width=15)
    file_input_label = Label(input_window, text="")
    slp_year_label = Label(input_window, text="Year for SLP calculation")
    excel_interval_entry = Entry(input_window, bd=5, width=3)
    start_label = Label(input_window, text="Start time")
    end_label = Label(input_window, text="End time")
    graphics_query_label = Label(input_window, text="Visualisation:")
    excel_query_label = Label(input_window, text="Cache interval:")
    excel_suffix_label = Label(input_window, text="hours")
    empty_label = Label(input_window, text="")
    empty_label2 = Label(input_window, text="")
    var1 = IntVar()
    c1 = Checkbutton(input_window, text ='Yes',variable = var1, onvalue = 1, offvalue = 0, command = "")
    instruction_label = Label(input_window, text = "Input File")
    ml_status_descr = Label(input_window, text = "ML model status:")
    ml_status_label = Label(input_window, text = "Please choose an input file.")
    ml_date_range_descr = Label(input_window, text = "Last change of ML models:")
    ml_date_range_label = Label(input_window, text = "Please choose an input file.")
    ml_calc_button = Button(input_window, text="Calculate ML models", command=button_calc_ml_models)
    slp_current_year_descr = Label(input_window, text="Current SLP year:")
    slp_current_year_label = Label(input_window, text=slp_current_year)
    slp_current_file_descr = Label(input_window, text="Current SLP input file:")
    slp_current_file_label = Label(input_window, text=slp_current_file)


    realtime_query_label = Label(input_window, text = "Real-time simulation (tbd):")
    var2 = IntVar()
    c2 = Checkbutton(input_window, text ='Yes',variable = var2, onvalue = 1, offvalue = 0, command = "")

    # BOOKMARK: Change default times here
    input_box_start.insert(0, "29.12.2022 00:05")
    input_box_end.insert(0, "30.12.2022 00:05")

    # Now we add the components to our window 
    # in the desired order.

    instruction_label.grid(row = 0, column = 0, sticky='e')
    input_file_name.grid(row = 0, column = 1)
    select_button.grid(row = 0, column = 2, sticky='w')

    slp_year_label.grid(row = 1, column = 0, sticky='e')
    input_year_slp.grid(row = 1, column = 1)
    data_prep_button.grid(row = 1, column = 2, sticky='w')

    slp_current_year_descr.grid(row = 2, column = 0, sticky='e')
    slp_current_year_label.grid(row = 2, column = 1)
    slp_current_file_descr.grid(row = 3, column = 0, sticky='e')
    slp_current_file_label.grid(row = 3, column = 1)


    ml_status_descr.grid(row=4, column=0, sticky='e')
    ml_status_label.grid(row=4, column=1)
    ml_calc_button.grid(row=4, column=2, sticky='w')
    ml_date_range_descr.grid(row=5, column=0, sticky='e')
    ml_date_range_label.grid(row=5, column=1)

    start_label.grid(row=6, column=0, sticky='e')
    input_box_start.grid(row=6, column=1)
    end_label.grid(row=7, column=0, sticky='e')
    input_box_end.grid(row=7, column=1)
    graphics_query_label.grid(row=8, column=0, sticky='e')
    c1.grid(row=8, column=1)
    realtime_query_label.grid(row=9, column=0, sticky='e')
    c2.grid(row=9, column=1)
    excel_query_label.grid(row=10, column=0, sticky='e')
    excel_interval_entry.grid(row=10, column=1)
    excel_suffix_label.grid(row=10, column=2, sticky='w')
    empty_label.grid(row=11, column=0, sticky='e')
    change_button.grid(row=12, column=1)
    empty_label2.grid(row=13, column=0)

    # Watch for changes in input year
    input_year_slp.bind('<KeyRelease>', on_year_slp_change)

    # Change the formatting to add some padding.
    for child in input_window.winfo_children():
        child.grid_configure(padx=5, pady=5)
    
    # Make first column left-bound
    input_window.grid_columnconfigure(0, weight=1,)


    # Set the focus to the select button
    select_button.focus()

    input_window.mainloop()
