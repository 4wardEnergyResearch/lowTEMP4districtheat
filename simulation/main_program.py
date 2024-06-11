# -*- coding: utf-8 -*-
"""
Author: 4wardEnergy Research GmbH
Date: 2024-05-14
Version: 1.0

Contains the main simulation loop.
Functions to read data, balance data, fill gaps, simulate the network 
hydraulically and thermally, calculate network losses, and save the data.

Functions:
- print_green: Prints text in green color.
- main_program: Executes the main program.
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

import numpy as np
import math
from openpyxl.styles import Font
from openpyxl.chart import ScatterChart, Reference, Series
from datetime import timedelta, datetime
import tkinter as tk
from pandas import DataFrame
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from tkinter import *
from tkinter.filedialog import askopenfilename
import importlib
import termcolor
import os

from simulation import output
from simulation import fcns_read
from simulation import fcns_balance
from simulation import fcns_gaps
from simulation import Auxiliary_functions
from simulation import hydraulic_equ_system
from simulation import thermal_equ_system


import options as options
importlib.reload(options)


def print_green(text):
    print(termcolor.colored(text, 'green'))

def main_program(var_phy, var_H2O, var_misc, var_sim, plots, file_input, \
                 excel_save_time, var_unused, var_gaps):
    """Executes the main program:
    Reads the data, balances the data, fills gaps in the data, simulates the
    network hydraulically and thermally, calculates the network losses and
    saves the data.

    :param var_phy: Physical variables
    :type var_phy: var_phy obj.
    
    :param var_H2O: Variables concerning water
    :type var_H2O: var_H2O obj.
    
    :param var_misc: Miscellaneous variables
    :type var_misc: var_misc obj.
    
    :param var_sim: Variables concerning the simulation
    :type var_sim: var_sim obj.
    
    :param plots: Variables concerning the plots
    :type plots: plots obj.
    
    :param file_input: Path of the input file
    :type file_input: str
    
    :param var_unused: Unused variables (Legacy)
    :type var_unused: var_unused obj.

    :param var_gaps: Variables concerning the gapfilling.
    :type var_gaps: var_gaps obj.
    """    

    ###########################################################################
    # READ DATA ###############################################################
    ###########################################################################

    fileXLSX, fileXLSX_name, node, line = fcns_read.read_data(var_misc, \
        var_sim, var_gaps, file_input)

###############################################################################
# INITIALIZATION FOR THE TRANSIENT SIMULATION #################################
###############################################################################
    total_time = 0
    # INITIALIZE ARRAYS
    line.n, line.dx, line.x,  \
        line.t_forerun, line.t_return, line.t_forerun_trans, line.t_return_trans, \
        line.dTdt, line.m_int_forerun, line.m_int_return, line.t_int_in, line.t_int_in_check, \
        line.t_int_out, line.t_int_out_check,  line.m_int_forerun_trans, line.m_int_return_trans = ([] for i in range(16))

    var_sim.temp_soil, var_sim.temp_soil_start = [], []

    adding_1, adding_2 = [], []

    node.m_ext_forerun_trans, node.m_ext_return_trans, node.p_forerun_trans, node.p_return_trans, node.t_forerun_trans, node.t_return_trans = ([] for i in range(6))

    # SOIL TEMPERATURE AT THE START OF THE SIMULATION #########################
    var_sim.temp_soil_start = Auxiliary_functions.soil_temp(var_sim.time_sim_start)
    


    for x_line in range (0, line.nbr_matrix.shape[0]):
        
        adding_1 = np.array([math.ceil(line.l[x_line]*var_sim.n_m)])
        adding_2 = np.array(line.l[x_line]/adding_1)
        line.n = np.append(line.n, adding_1, axis = 0)
        line.dx = np.append(line.dx, adding_2, axis = 0)
        line.x.append(np.linspace(line.dx[x_line]/2, line.l[x_line]-line.dx[x_line]/2, int(line.n[x_line])))
        line.t_forerun.append(np.ones(int(line.n[x_line]))*var_sim.temp_soil_start)
        line.t_return.append(np.ones(int(line.n[x_line]))*var_sim.temp_soil_start)
        line.t_forerun_trans.append(np.array([]))
        line.t_return_trans.append(np.array([]))
        line.dTdt.append(np.zeros(int(line.n[x_line])))

    del adding_1, adding_2

    # Since all internal mass flows and almost all pressures / temperatures
    # are calculated rather than read, arrays are created only for the values
    # and for verification, respectively
    line.m_int_forerun = np.ones(line.nbr_matrix.shape[0])
    line.m_int_return = np.ones(line.nbr_matrix.shape[0])

    line.t_int_in =          [None]*line.nbr_matrix.shape[0]
    line.t_int_in_check =    [None]*line.nbr_matrix.shape[0]
    line.t_int_out =         [None]*line.nbr_matrix.shape[0]
    line.t_int_out_check =   [None]*line.nbr_matrix.shape[0]

    var_sim.time_sim = var_sim.time_sim_start
    var_sim.cntr_time_hyd = 0
    var_sim.cntr_time_therm_forerun = 0
    var_sim.cntr_time_therm_return = 0
    var_sim.cntr = 0

    # INITIALIZE BALANCE OBJECT
    class balance:
        pass

    balance = balance()

    # Sanity check for p_ref:
    if len(node.p_ref) == 0:
        print("No node has been specified for the reference pressure. Please specify exactly one node.")
        exit()
    if len(node.p_ref) > 1:
        print("More than one node has been specified for the reference pressure. Please specify exactly one node.")
        exit()

    # Calculate time step interval for thermal solution plotting
    plots.thermal_update_time_steps = np.round(plots.update_interval / var_sim.delta_time_therm)

    # If the plots output_dir does not exist, create it
    if plots.show_plot == "yes":
        if not os.path.exists(plots.output_dir):
            os.makedirs(plots.output_dir)


# BOOKMARK: Start of main loop
    while var_sim.time_sim <= var_sim.time_sim_end:
        print(str(var_sim.time_sim))
        loop_start = datetime.now()

        # [ ] Check for real-time data; if so, read data from the last 15 minutes. This is to be implemented in DDM.
        
        #######################################################################
        # BALANCING ###########################################################
        #######################################################################

        balance = fcns_balance.balance_meas(var_misc, var_sim, fileXLSX, \
            fileXLSX_name, node, balance)

        #######################################################################
        # GAPFILLING AND BALANCING ############################################
        #######################################################################

        fcns_gaps.gaps(fileXLSX, fileXLSX_name, node, var_sim, balance, var_gaps)

        fcns_balance.balance_sim(var_misc, var_sim, fileXLSX, \
            fileXLSX_name, node, balance)        
        
        #######################################################################
        # INITIALIZE ARRAYS FOR THE TRANSIENT SIMULATION ######################
        #######################################################################
        # SOIL TEMPERATURE
        var_sim.temp_soil = np.append(var_sim.temp_soil, np.array([Auxiliary_functions.soil_temp(var_sim.time_sim)]), axis = 0)

        # EXTERNAL MASS FLOWS
        node.m_ext_forerun, node.m_ext_forerun_check, node.m_ext_return, node.m_ext_return_check = ([] for i in range(4))

        for x_node in range (0, node.nbr_matrix.shape[0]):
            # DISTRIBUTOR
            if node.distrib[x_node] == "x":
                node.m_ext_forerun =        np.append(node.m_ext_forerun,       np.array([0]), axis = 0)
                node.m_ext_forerun_check =  np.append(node.m_ext_forerun_check, np.array([0]), axis = 0)
                node.m_ext_return =         np.append(node.m_ext_return,        np.array([0]), axis = 0)
                node.m_ext_return_check =   np.append(node.m_ext_return_check,  np.array([0]), axis = 0)
            # FEEDER (except for the feeder where the reference pressure is given)
            elif (node.feed_in[x_node]) and (x_node != node.p_ref[0]):
                m_ext_current = node.V_dot_feed[x_node][var_sim.cntr_time_hyd] * var_H2O.rho / 1000
                node.m_ext_forerun =        np.append(node.m_ext_forerun,       np.array([m_ext_current]), axis = 0)
                node.m_ext_forerun_check =  np.append(node.m_ext_forerun_check, np.array([m_ext_current]), axis = 0)
                node.m_ext_return =         np.append(node.m_ext_return,        np.array([-m_ext_current]), axis = 0)
                node.m_ext_return_check =   np.append(node.m_ext_return_check,  np.array([-m_ext_current]), axis = 0)
            # NO MEASUREMENTS (or feeder with reference pressure)
            elif type(node.V_dot_sim[x_node]) is not np.ndarray:
                node.m_ext_forerun =        np.append(node.m_ext_forerun,       np.array([None]), axis = 0)
                node.m_ext_forerun_check =  np.append(node.m_ext_forerun_check, np.array([None]), axis = 0)
                node.m_ext_return =         np.append(node.m_ext_return,        np.array([None]), axis = 0)
                node.m_ext_return_check =   np.append(node.m_ext_return_check,  np.array([None]), axis = 0)
            # CONSUMER
            else:
                m_ext_current = node.V_dot_sim[x_node][var_sim.cntr_time_hyd] * var_H2O.rho / 1000
                node.m_ext_forerun =        np.append(node.m_ext_forerun,       np.array([-m_ext_current]), axis = 0)
                node.m_ext_forerun_check =  np.append(node.m_ext_forerun_check, np.array([-m_ext_current]), axis = 0)
                node.m_ext_return =         np.append(node.m_ext_return,        np.array([m_ext_current]), axis = 0)
                node.m_ext_return_check =   np.append(node.m_ext_return_check,  np.array([m_ext_current]), axis = 0)

        # PRESSURES
        node.p_forerun, node.p_ref_forerun, node.p_return, node.p_ref_return = [], [], [], []
        for x_node in range (0, node.nbr_matrix.shape[0]):
            cntr_p_ref = 0
            for x_p_ref_node in range (0, len(node.p_ref)):
                if x_node == node.p_ref[x_p_ref_node]:
                    node.p_forerun = np.append(node.p_forerun, np.array([node.p_flow_feed[x_node][var_sim.cntr_time_hyd]]))
                    node.p_ref_forerun = np.append(node.p_ref_forerun, np.array([node.p_flow_feed[x_node][var_sim.cntr_time_hyd]]))
                    node.p_return = np.append(node.p_return, np.array([node.p_ret_feed[x_node][var_sim.cntr_time_hyd]]))
                    node.p_ref_return = np.append(node.p_ref_return, np.array([node.p_ret_feed[x_node][var_sim.cntr_time_hyd]]))
                    cntr_p_ref = cntr_p_ref+1
            if cntr_p_ref == 0:
                node.p_forerun = np.append(node.p_forerun, np.array([None]))
                node.p_ref_forerun = np.append(node.p_ref_forerun, np.array([None]))
                node.p_return = np.append(node.p_return, np.array([None]))
                node.p_ref_return = np.append(node.p_ref_return, np.array([None]))

# BOOKMARK: Hydr. and thermal calculations
        #######################################################################
        # HYDRAULIC CALCULATION ###############################################
        #######################################################################
        
        # CREATE COUPLING MATRIX
        if var_sim.cntr_time_hyd == 0:
            Auxiliary_functions.make_matrix_coupl(line, node, var_sim)

        # SET START VALUES FOR THE SOLVER
        Auxiliary_functions.number_unknowns(line, node, var_sim, var_H2O)

        # SOLVE THE HYDRAULIC EQUATION SYSTEM
        hydraulic_equ_system.solve_network_hydr(var_phy, var_H2O, var_sim, var_misc, line, node, plots)

        #######################################################################
        # THERMAL TRANSIENT CALCULATION #######################################
        #######################################################################
        # SOLVE THE THERMAL EQUATION SYSTEM
        thermal_equ_system.solve_network_therm(var_phy, var_H2O, var_sim, var_misc, line, node, plots)

        var_sim.cntr_time_hyd = var_sim.cntr_time_hyd+1
        
        
        #######################################################################
        # NETWORK LOSSES ######################################################
        #######################################################################
        
        line.Q_dot_forerun_line_loss = []
        line.Q_dot_return_line_loss = []
        line.q_dot_forerun_line_loss = []
        line.q_dot_return_line_loss = []
        
        # BACKUP
        if excel_save_time != 0:
            if ((var_sim.time_sim-var_sim.time_sim_start).total_seconds()/3600)%excel_save_time == 0 and var_sim.time_sim != var_sim.time_sim_start:
                print("Caching: " + var_sim.time_sim.strftime(format = "%Y-%m-%d %H:%M"))
                var_sim.time_calc = var_sim.time_sim_start
                var_sim.cntr = 0
                while var_sim.time_calc <= var_sim.time_sim:
                    Q_dot_forerun_line_loss_add, Q_dot_return_line_loss_add = [], []
                    q_dot_forerun_line_loss_add, q_dot_return_line_loss_add = [], []
                    for x_line in range(0, line.node_start.shape[0]):
                        xx = int(var_sim.cntr*var_sim.delta_time_hyd*60/var_sim.delta_time_therm)
                        
                        ###########################################################
                        # LINE LOSSES FLOW ########################################
                        ###########################################################
                        # Array for the arithmetic mean value formation of the individual temperatures of the sectors of the line over time
                        t_line__forerun_array = line.t_forerun_trans[x_line][xx:int(xx+var_sim.delta_time_hyd*60/var_sim.delta_time_therm),:]     
                        t_line_forerun_mean = np.mean(t_line__forerun_array)
                        #Q_dot_forerun_line_loss_add = np.append(Q_dot_forerun_line_loss_add, np.array([(t_line_forerun_mean-temp_soil_add)*line.htc[x_line]*line.l[x_line]*0.001]), axis = 0)
                        Q_dot_forerun_line_loss_add = np.append(Q_dot_forerun_line_loss_add, np.array([(t_line_forerun_mean-var_sim.temp_soil[var_sim.cntr])*line.htc[x_line]*line.l[x_line]*0.001]), axis = 0)
                        q_dot_forerun_line_loss_add = np.append(q_dot_forerun_line_loss_add, line.htc[x_line]/(var_H2O.c_p*abs(line.m_int_forerun_trans[x_line][var_sim.cntr])), axis = 0)

                        ###########################################################
                        # LINE LOSSES RETURN ######################################
                        ###########################################################
                        t_line__return_array = line.t_return_trans[x_line][xx:int(xx+var_sim.delta_time_hyd*60/var_sim.delta_time_therm),:]
                        t_line_return_mean = np.mean(t_line__return_array)
                        #Q_dot_return_line_loss_add = np.append(line.Q_dot_return_line_loss_add, np.array([(t_line_return_mean-temp_soil_add)*line.htc[x_line]*line.l[x_line]*0.001]), axis = 0)
                        Q_dot_return_line_loss_add = np.append(Q_dot_return_line_loss_add, np.array([(t_line_return_mean-var_sim.temp_soil[var_sim.cntr])*line.htc[x_line]*line.l[x_line]*0.001]), axis = 0)
                        q_dot_return_line_loss_add = np.append(q_dot_return_line_loss_add, line.htc[x_line]/(var_H2O.c_p*abs(line.m_int_return_trans[x_line][var_sim.cntr])), axis = 0)

                        #print(t_forerun_trans[x_line][xx][0])
                        #print(t_forerun_trans[x_line][xx][-1])
                        #print(line.m_int_forerun_trans[x_line][cntr])
                        #print(Q_dot_forerun_line_loss_add)
                        #Q_dot_forerun_line_loss_add = np.append(Q_dot_forerun_line_loss_add, np.array((t_forerun_trans[x_line][xx][0]-t_forerun_trans[x_line][xx][-1])*var_H2O.c_p*line.m_int_forerun_trans[x_line][cntr]*0.001), axis = 0)
                        #Q_dot_return_line_loss_add = np.append(line.Q_dot_return_line_loss_add, np.array((t_return_trans[x_line][xx][-1]-t_return_trans[x_line][xx][0])*var_H2O.c_p*line.m_int_return_trans[x_line][cntr]*0.001), axis = 0)
                    line.Q_dot_forerun_line_loss.append(Q_dot_forerun_line_loss_add)
                    line.Q_dot_return_line_loss.append(Q_dot_return_line_loss_add)
                    line.q_dot_forerun_line_loss.append(q_dot_forerun_line_loss_add)
                    line.q_dot_return_line_loss.append(q_dot_return_line_loss_add)
                    var_sim.time_calc = var_sim.time_calc+timedelta(seconds = 60*var_sim.delta_time_hyd)
                    var_sim.cntr += 1
                
                #cntr = cntr-1
                output.Save_Excel(var_H2O, var_sim, var_misc, var_unused, line, node, fileXLSX, fileXLSX_name)
            else:
                pass
            
        var_sim.time_sim = var_sim.time_sim+timedelta(minutes = var_sim.delta_time_hyd)
        time_spent_in_loop = datetime.now()-loop_start
        total_time += time_spent_in_loop.total_seconds()

        print_green("Hydraulic timestep completed. Calculation time: " + str(round(time_spent_in_loop.total_seconds(), 1)) + " s")
    
    ###########################################################################
    # NETWORK LOSSES ##########################################################
    ###########################################################################

    var_sim.time_calc = var_sim.time_sim_start
    var_sim.time_sim = var_sim.time_sim-timedelta(minutes = var_sim.delta_time_hyd)
    var_sim.cntr = 0
    while var_sim.time_calc <= var_sim.time_sim_end:
        Q_dot_forerun_line_loss_add, Q_dot_return_line_loss_add = [], []
        for x_line in range(0, line.node_start.shape[0]):
            xx = int(var_sim.cntr*var_sim.delta_time_hyd*60/var_sim.delta_time_therm)
            Q_dot_forerun_line_loss_add = np.append(Q_dot_forerun_line_loss_add, np.array((line.t_forerun_trans[x_line][xx][0]-line.t_forerun_trans[x_line][xx][-1])*var_H2O.c_p*line.m_int_forerun_trans[x_line][var_sim.cntr]*0.001), axis = 0)
            Q_dot_return_line_loss_add = np.append(Q_dot_return_line_loss_add, np.array((line.t_return_trans[x_line][xx][-1]-line.t_return_trans[x_line][xx][0])*var_H2O.c_p*line.m_int_return_trans[x_line][var_sim.cntr]*0.001), axis = 0)
        line.Q_dot_forerun_line_loss.append(Q_dot_forerun_line_loss_add)
        line.Q_dot_return_line_loss.append(Q_dot_return_line_loss_add)
        try:
            line.q_dot_forerun_line_loss.append(q_dot_forerun_line_loss_add)
        except:
            pass
        try:
            line.q_dot_return_line_loss.append(q_dot_return_line_loss_add)
        except:
            pass
        var_sim.time_calc = var_sim.time_calc+timedelta(seconds = 60*var_sim.delta_time_hyd)
        var_sim.cntr += 1

    #cntr -= 1
    print("Saving results: " + var_sim.time_sim.strftime(format = "%Y-%m-%d %H:%M"))
    output.Save_Excel(var_H2O, var_sim, var_misc, var_unused, line, node, fileXLSX, fileXLSX_name)
    print_green("Mean time elapsed per hydraulic timestep: " + str(round(total_time/var_sim.cntr_time_hyd+1, 1)) + " s")